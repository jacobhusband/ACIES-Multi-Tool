// ===================== CONFIGURATION & CONSTANTS =====================
const STATUS_CANON = [
  "Waiting",
  "Working",
  "Pending Review",
  "Complete",
  "Delivered",
];
const STATUS_PRIORITY = [
  "Delivered",
  "Complete",
  "Pending Review",
  "Working",
  "Waiting",
];
const LABEL_TO_KEY = {
  Waiting: "waiting",
  Working: "working",
  "Pending Review": "pendingReview",
  Complete: "complete",
  Delivered: "delivered",
};
const KEY_TO_LABEL = {
  waiting: "Waiting",
  working: "Working",
  pendingReview: "Pending Review",
  complete: "Complete",
  delivered: "Delivered",
};
const HELP_LINKS = {
  main: "https://brainy-seahorse-3c5.notion.site/ACIES-Desktop-Application-2b03fdbb662c80afa61af555bddc9e61?pvs=74",
  projects:
    "https://brainy-seahorse-3c5.notion.site/Projects-2b13fdbb662c803b9898d86ec9389e41",
  notes:
    "https://brainy-seahorse-3c5.notion.site/Notes-2b13fdbb662c80dd86c8e9d3f0d65081",
  tools:
    "https://brainy-seahorse-3c5.notion.site/Tools-2b13fdbb662c80afbab9d68204f9cd23",
  plugins:
    "https://brainy-seahorse-3c5.notion.site/Plugins-2b13fdbb662c801abfe9cc46927eb73a",
  timesheets:
    "https://brainy-seahorse-3c5.notion.site/ACIES-Desktop-Application-2b03fdbb662c80afa61af555bddc9e61?pvs=74",
};
const THEME_STORAGE_KEY = "acies-theme";
const LIGHTING_SCHEDULE_FIELDS = [
  "mark",
  "description",
  "manufacturer",
  "modelNumber",
  "mounting",
  "volts",
  "watts",
  "notes",
];
const LIGHTING_SCHEDULE_DEFAULT_GENERAL_NOTES = [
  "A.  VERIFY ALL CEILING TYPES PRIOR TO ORDERING FIXTURES.",
  "B.  VERIFY ALL OPERATING VOLTAGE PRIOR TO ORDERING FIXTURES.",
  "C.  COORDINATE THE HEIGHT OF ALL SUSPENDED FIXTURES WITH THE OWNER, ARCHITECT AND ENGINEER PRIOR TO INSTALLATION.",
  "D.  ALL LAMPS SHALL HAVE A COLOR TEMPERATURE OF 3500 DEG. KELVIN AND A CRI OF 85 UNLESS SPECIFICALLY NOTED.",
  'E.  FIXTURES DESIGNATED WITH A "1/2 SHADE" OR "FULL SHADE" AND ALL EXIT SIGNS SHALL BE CIRCUITED TO THE CENTRAL EMERGENCY BATTERY INVERTER OR PROVIDED WITH INTEGRAL BATTERY BACKUP AS REQUIRED. EM BACKUP POWER SHALL BE SUITABLE TO PROVIDE FULL POWER TO FIXTURES FOR A MINIMUM OF 90 MINUTES.',
].join("\n");
const LIGHTING_SCHEDULE_DEFAULT_NOTES = [
  "1.  CONFIRM THE DRIVER TYPE WITH VENDOR. LIGHT DRIVER SHALL BE COMPATIBLE WITH THE DIMMER. REFER TO THE SENSOR SCHEDULE FOR DETAILS.",
  "2.  COORDINATE THE FINAL MANUFACTURER AND MODEL WITH ARCHITECT.",
].join("\n");
const STAR_ICON_PATH =
  "M12 17.27L18.18 21l-1.64-7.03L22 9.24l-7.19-.61L12 2 9.19 8.63 2 9.24l5.46 4.73L5.82 21z";
const EYE_ICON_PATH =
  "M12 4.5C7 4.5 2.73 7.61 1 12c1.73 4.39 6 7.5 11 7.5s9.27-3.11 11-7.5C21.27 7.61 17 4.5 12 4.5zm0 12.5a5 5 0 1 1 0-10 5 5 0 0 1 0 10zm0-8a3 3 0 1 0 0 6 3 3 0 0 0 0-6z";
const PIN_ICON_PATH =
  "M12 2C8.13 2 5 5.13 5 9c0 2.76 1.87 5.08 4.42 5.76L10 22l2-3 2 3-.42-7.24C16.13 14.08 18 11.76 18 9c0-3.87-3.13-7-7-7zm0 8.5A2.5 2.5 0 1 1 12 5a2.5 2.5 0 0 1 0 5z";
const PENCIL_ICON_PATH =
  "M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zm17.71-10.21c.39-.39.39-1.02 0-1.41l-2.34-2.34a1 1 0 0 0-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z";
const COPY_ICON_PATH =
  "M16 1H4c-1.1 0-2 .9-2 2v14h2V3h12V1zm3 4H8c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h11c1.1 0 2-.9 2-2V7c0-1.1-.9-2-2-2zm0 16H8V7h11v14z";
const TRASH_ICON_PATH =
  "M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zm13-15h-3.5l-1-1h-5l-1 1H5v2h14V4z";

// Timesheet Constants
const DISCIPLINE_TO_FUNCTION = {
  Electrical: "E",
  Mechanical: "M",
  Plumbing: "P",
};
const MAX_HOURS_PER_DAY = 8;
const WFH_SUFFIX = " - WFH";
const TIMESHEET_SUMMARY_DELIVERABLE_ID = "__summary__";
const WEEKLY_MEETING_NAME = "workload meeting";
const WEEKLY_MEETING_DEFAULT_HOURS = 1;

// Cache for loaded descriptions to prevent re-fetching
const DESCRIPTION_CACHE = {};
let bundlesPrefetchStarted = false;
let bundlesCache = null;
let bundlesCacheKey = "";
let bundlesLastRenderKey = "";
let bundlesLoadingPromise = null;
let bundlesNeedsRender = false;
let bundlesLastCheckAt = 0;
let bundlesUpdateCheckInFlight = false;

// ===================== UTILITIES & HELPERS =====================

function el(tag, props = {}, children = []) {
  const n = document.createElement(tag);
  Object.entries(props).forEach(([k, v]) => {
    if (k.startsWith("aria-") || k.startsWith("data-")) {
      if (v != null) n.setAttribute(k, String(v));
    } else {
      n[k] = v;
    }
  });
  children.forEach((c) => {
    if (typeof c === "string") n.appendChild(document.createTextNode(c));
    else n.appendChild(c);
  });
  return n;
}

// ===================== TIMESHEET UTILITIES =====================

function getWeekStartDate(date) {
  const d = new Date(date);
  const day = d.getDay(); // 0 = Sunday
  d.setDate(d.getDate() - day);
  d.setHours(0, 0, 0, 0);
  return d;
}

function formatWeekKey(date) {
  const d = getWeekStartDate(date);
  return d.toISOString().split("T")[0];
}

function formatWeekDisplay(date) {
  const start = getWeekStartDate(date);
  const end = new Date(start);
  end.setDate(start.getDate() + 6);
  const options = { month: "short", day: "numeric" };
  const startStr = start.toLocaleDateString(undefined, options);
  const endStr = end.toLocaleDateString(undefined, {
    ...options,
    year: "numeric",
  });
  return `${startStr} - ${endStr}`;
}

function normalizeDisciplineList(value) {
  if (Array.isArray(value)) return value;
  if (typeof value === "string") {
    const parts = value
      .split(/[,/;]+/)
      .map((item) => item.trim())
      .filter(Boolean);
    return parts.length ? parts : ["Electrical"];
  }
  return ["Electrical"];
}

function getDisciplineFunction() {
  const disciplines = normalizeDisciplineList(userSettings.discipline);
  const funcs = disciplines
    .map((d) => DISCIPLINE_TO_FUNCTION[d] || "")
    .filter(Boolean);
  return funcs.join("/") || "E";
}

function getDefaultPmInitials() {
  return String(userSettings.defaultPmInitials || "").trim().toUpperCase();
}

function createTimesheetEntryId() {
  return "ts_" + Math.random().toString(36).substr(2, 9);
}

function getTimesheetEntryDescription(entry, fallback = "") {
  const desc = entry?.serviceDescription || entry?.deliverableName || fallback || "";
  return String(desc).trim();
}

function getTimesheetTaskNumberForDeliverable(name) {
  const text = String(name || "").toLowerCase();
  if (text.includes("rfi") || text.includes("submittal")) return "0.CA";
  if (text.includes("asr")) return "0.05";
  if (text.includes("survey")) return "0.SV";
  return "0.00";
}

function hasWfhSuffix(value) {
  const trimmed = String(value || "").trimEnd();
  const upper = trimmed.toUpperCase();
  return upper === "WFH" || upper.endsWith(WFH_SUFFIX);
}

function stripWfhSuffix(value) {
  const trimmed = String(value || "").trimEnd();
  const upper = trimmed.toUpperCase();
  if (upper === "WFH") return "";
  if (upper.endsWith(WFH_SUFFIX)) {
    return trimmed.slice(0, -WFH_SUFFIX.length).trimEnd();
  }
  return trimmed;
}

function appendWfhSuffix(value) {
  const trimmed = String(value || "").trimEnd();
  if (!trimmed) return "WFH";
  return hasWfhSuffix(trimmed) ? trimmed : `${trimmed}${WFH_SUFFIX}`;
}

function normalizeDeliverableNames(deliverables) {
  const names = [];
  const seen = new Set();
  (deliverables || []).forEach((deliverable) => {
    const raw =
      typeof deliverable === "string" ? deliverable : deliverable?.name || "";
    String(raw || "")
      .split(",")
      .map((item) => item.trim())
      .filter(Boolean)
      .forEach((name) => {
        const key = name.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        names.push(name);
      });
  });
  return names;
}

function parseDeliverableList(description) {
  return String(description || "")
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
}

function mergeDeliverableDescription(existingDescription, deliverableNames) {
  const baseList = parseDeliverableList(existingDescription);
  const seen = new Set(baseList.map((name) => name.toLowerCase()));
  const merged = baseList.slice();
  deliverableNames.forEach((name) => {
    const key = String(name || "").toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    merged.push(name);
  });
  return merged.join(", ");
}

function getMissingDeliverables(existingDescription, deliverableNames) {
  const baseList = parseDeliverableList(existingDescription);
  const seen = new Set(baseList.map((name) => name.toLowerCase()));
  return deliverableNames.filter(
    (name) => !seen.has(String(name || "").toLowerCase())
  );
}

function getTimesheetProjectMatchKey(project, fallbackIndex = null) {
  const id = String(project?.id || "").trim();
  if (id) return `id:${id.toLowerCase()}`;
  const name = String(project?.nick || project?.name || "").trim();
  if (name) return `name:${name.toLowerCase()}`;
  if (fallbackIndex != null) return `project:${fallbackIndex}`;
  return "";
}

function getTimesheetEntryMatchKey(entry) {
  const projectId = String(entry?.projectId || "").trim();
  if (projectId) return `id:${projectId.toLowerCase()}`;
  const projectName = String(entry?.projectName || "").trim();
  if (projectName) return `name:${projectName.toLowerCase()}`;
  return "";
}

function isProjectSummaryEntry(entry) {
  return (
    entry?.deliverableId === TIMESHEET_SUMMARY_DELIVERABLE_ID ||
    entry?.isProjectSummary
  );
}

function getSummaryEntryStatus(summaryEntries, deliverableNames) {
  const status = {
    count: 0,
    hasOffice: false,
    hasWfh: false,
    needsUpdate: false,
  };
  (summaryEntries || []).forEach((entry) => {
    status.count += 1;
    const desc = getTimesheetEntryDescription(entry);
    const isWfh = hasWfhSuffix(desc);
    if (isWfh) {
      status.hasWfh = true;
    } else {
      status.hasOffice = true;
    }
    const baseDesc = stripWfhSuffix(desc);
    if (getMissingDeliverables(baseDesc, deliverableNames).length) {
      status.needsUpdate = true;
    }
  });
  return status;
}

function getProjectsForWeek(weekStart) {
  const weekEnd = new Date(weekStart);
  weekEnd.setDate(weekStart.getDate() + 13); // This week + next week

  const projectsWithDeliverables = [];

  db.forEach((project) => {
    const deliverables = getProjectDeliverables(project);
    deliverables.forEach((deliverable) => {
      const due = parseDueStr(deliverable?.due);
      if (due && due >= weekStart && due <= weekEnd) {
        projectsWithDeliverables.push({
          project,
          deliverable,
          dueDate: due,
        });
      }
    });
  });

  // Sort: this week first, then by due date
  const thisWeekEnd = new Date(weekStart);
  thisWeekEnd.setDate(weekStart.getDate() + 6);

  return projectsWithDeliverables.sort((a, b) => {
    const aThisWeek = a.dueDate <= thisWeekEnd;
    const bThisWeek = b.dueDate <= thisWeekEnd;
    if (aThisWeek && !bThisWeek) return -1;
    if (!aThisWeek && bThisWeek) return 1;
    return a.dueDate - b.dueDate;
  });
}

function getWeekEntries(weekKey) {
  return timesheetDb.weeks[weekKey]?.entries || [];
}

function setWeekEntries(weekKey, entries) {
  if (!timesheetDb.weeks[weekKey]) {
    timesheetDb.weeks[weekKey] = { entries: [], notes: "" };
  }
  timesheetDb.weeks[weekKey].entries = entries;
}

let timesheetDragEntryId = null;

function ensureTimesheetEntryIds(entries) {
  let updated = false;
  entries.forEach((entry) => {
    if (!entry.id) {
      entry.id = createTimesheetEntryId();
      updated = true;
    }
  });
  if (updated) saveTimesheets();
}

function moveTimesheetEntry(entries, fromIndex, toIndex) {
  if (fromIndex === toIndex) return entries;
  const boundedTo = Math.max(0, Math.min(toIndex, entries.length));
  const [moved] = entries.splice(fromIndex, 1);
  entries.splice(boundedTo, 0, moved);
  return entries;
}

function clearTimesheetDragStyles() {
  document.querySelectorAll(".ts-entry-row").forEach((row) => {
    row.classList.remove("ts-drop-before", "ts-drop-after", "ts-dragging");
    delete row.dataset.dropPosition;
  });
}

// ===================== TIMESHEET I/O =====================

async function loadTimesheets() {
  try {
    const data = await window.pywebview.api.get_timesheets();
    timesheetDb = data || { weeks: {}, lastModified: null };
    return timesheetDb;
  } catch (e) {
    console.warn("Failed to load timesheets:", e);
    return { weeks: {}, lastModified: null };
  }
}

async function saveTimesheets() {
  try {
    timesheetDb.lastModified = new Date().toISOString();
    const response = await window.pywebview.api.save_timesheets(timesheetDb);
    if (response.status !== "success") throw new Error(response.message);
  } catch (e) {
    console.warn("Failed to save timesheets:", e);
    toast("Failed to save timesheet data.");
  }
}

// ===================== TEMPLATES I/O =====================

async function loadTemplates() {
  try {
    const data = await window.pywebview.api.get_templates();
    templatesDb = data || { templates: [], defaultTemplatesInstalled: false, lastModified: null };
    return templatesDb;
  } catch (e) {
    console.warn("Failed to load templates:", e);
    return { templates: [], defaultTemplatesInstalled: false, lastModified: null };
  }
}

async function saveTemplates() {
  try {
    templatesDb.lastModified = new Date().toISOString();
    const response = await window.pywebview.api.save_templates(templatesDb);
    if (response.status !== "success") throw new Error(response.message);
  } catch (e) {
    console.warn("Failed to save templates:", e);
    toast("Failed to save templates.");
  }
}

// ===================== TEMPLATES RENDERING =====================

function renderTemplates() {
  const container = document.getElementById('templatesContainer');
  if (!container) return;

  const templates = templatesDb.templates || [];
  const emptyState = document.getElementById('templatesEmptyState');

  if (!templates.length) {
    container.innerHTML = '';
    if (emptyState) emptyState.style.display = 'block';
    return;
  }
  if (emptyState) emptyState.style.display = 'none';

  // Group templates by discipline
  const grouped = groupTemplatesByDiscipline(templates);

  container.innerHTML = '';

  // Get user disciplines for highlighting
  const userDisciplines = normalizeDisciplineList(userSettings.discipline);

  // Render General first, then user's disciplines, then others
  const disciplineOrder = ['General', ...userDisciplines.filter(d => d !== 'General')];
  const remainingDisciplines = DISCIPLINE_OPTIONS.filter(
    d => !disciplineOrder.includes(d) && grouped[d]?.length
  );
  const orderedDisciplines = [...disciplineOrder, ...remainingDisciplines];

  orderedDisciplines.forEach(discipline => {
    const disciplineTemplates = grouped[discipline];
    if (!disciplineTemplates?.length) return;

    const isUserDiscipline = discipline === 'General' || userDisciplines.includes(discipline);
    const section = createTemplateSection(discipline, disciplineTemplates, isUserDiscipline);
    container.appendChild(section);
  });
}

function groupTemplatesByDiscipline(templates) {
  const grouped = {};
  templates.forEach(template => {
    const discipline = template.discipline || 'General';
    if (!grouped[discipline]) grouped[discipline] = [];
    grouped[discipline].push(template);
  });
  return grouped;
}

function createTemplateSection(discipline, templates, isHighlighted) {
  const section = el('div', { className: 'templates-section' + (isHighlighted ? ' highlighted' : '') });

  const header = el('div', { className: 'templates-section-header' });
  header.appendChild(el('h4', { className: 'templates-section-title', textContent: discipline }));
  header.appendChild(el('p', {
    className: 'tiny muted templates-section-desc',
    textContent: `${templates.length} template${templates.length !== 1 ? 's' : ''}`
  }));
  section.appendChild(header);

  const grid = el('div', { className: 'templates-grid' });
  templates.forEach(template => {
    grid.appendChild(createTemplateCard(template));
  });
  section.appendChild(grid);

  return section;
}

function createTemplateCard(template) {
  const card = el('div', {
    className: 'template-card' + (template.isDefault ? ' default-template' : ''),
    'data-template-id': template.id
  });

  const icon = el('div', {
    className: 'template-icon',
    textContent: FILE_TYPE_ICONS[template.fileType] || '📄'
  });

  const content = el('div', { className: 'template-card-content' });

  const headerDiv = el('div', { className: 'template-card-header' });
  headerDiv.appendChild(el('span', {
    className: 'template-name',
    textContent: template.name
  }));
  if (template.isDefault) {
    headerDiv.appendChild(el('span', {
      className: 'template-badge default',
      textContent: 'Default'
    }));
  }
  content.appendChild(headerDiv);

  const meta = el('div', { className: 'template-meta' });
  meta.appendChild(el('span', {
    className: 'template-type',
    textContent: FILE_TYPE_LABELS[template.fileType] || template.fileType.toUpperCase()
  }));
  content.appendChild(meta);

  if (template.description) {
    content.appendChild(el('div', {
      className: 'template-description tiny muted',
      textContent: template.description
    }));
  }

  const actions = el('div', { className: 'template-actions' });

  const copyBtn = el('button', {
    className: 'btn btn-primary tiny',
    textContent: 'Copy to Folder',
    onclick: (e) => {
      e.stopPropagation();
      handleCopyTemplate(template.id);
    }
  });
  actions.appendChild(copyBtn);

  if (!template.isDefault) {
    const removeBtn = el('button', {
      className: 'btn ghost tiny text-danger',
      textContent: 'Remove',
      onclick: (e) => {
        e.stopPropagation();
        handleRemoveTemplate(template.id, template.name);
      }
    });
    actions.appendChild(removeBtn);
  }

  card.appendChild(icon);
  card.appendChild(content);
  card.appendChild(actions);

  return card;
}

// ===================== TEMPLATES ACTION HANDLERS =====================

async function handleAddTemplate() {
  const dlg = document.getElementById('addTemplateDlg');
  if (!dlg) return;

  // Reset form
  document.getElementById('tpl_name').value = '';
  document.getElementById('tpl_discipline').value = 'General';
  document.getElementById('tpl_path').value = '';
  document.getElementById('tpl_description').value = '';

  dlg.showModal();
}

async function handleBrowseTemplateFile() {
  try {
    const result = await window.pywebview.api.select_template_file();
    if (result.status === 'success' && result.path) {
      document.getElementById('tpl_path').value = result.path;

      // Auto-fill name from filename if empty
      const nameInput = document.getElementById('tpl_name');
      if (!nameInput.value) {
        const filename = result.path.split(/[\\/]/).pop();
        nameInput.value = filename.replace(/\.[^/.]+$/, ''); // Remove extension
      }
    }
  } catch (e) {
    toast('Error selecting file.');
  }
}

async function handleSaveNewTemplate() {
  const name = document.getElementById('tpl_name').value.trim();
  const discipline = document.getElementById('tpl_discipline').value;
  const sourcePath = document.getElementById('tpl_path').value.trim();
  const description = document.getElementById('tpl_description').value.trim();

  if (!name) {
    toast('Please enter a template name.');
    return;
  }
  if (!sourcePath) {
    toast('Please select a template file.');
    return;
  }

  try {
    const result = await window.pywebview.api.add_template(name, discipline, sourcePath, description);
    if (result.status === 'success') {
      toast('Template added successfully.');
      closeDlg('addTemplateDlg');
      await loadTemplates();
      renderTemplates();
    } else {
      toast(result.message || 'Failed to add template.');
    }
  } catch (e) {
    toast('Error adding template.');
  }
}

function getTemplateKey(template) {
  const name = String(template?.name || "").trim().toLowerCase();
  if (TEMPLATE_KEY_BY_NAME[name]) return TEMPLATE_KEY_BY_NAME[name];
  const sourceName = basename(template?.sourcePath || "").toLowerCase();
  if (sourceName.includes("narrative of changes")) return "narrative";
  if (sourceName.includes("plan check") || sourceName.includes("pcc")) return "planCheck";
  return null;
}

function getTemplateByKey(key) {
  const templates = templatesDb?.templates || [];
  return templates.find((template) => getTemplateKey(template) === key) || null;
}

function getPrimaryDisciplineLabel() {
  const disciplines = normalizeDisciplineList(userSettings.discipline);
  return disciplines[0] || "General";
}

function joinPath(base, child) {
  const baseStr = String(base || "").trim();
  const childStr = String(child || "").trim();
  if (!baseStr) return childStr;
  if (!childStr) return baseStr;
  const baseClean = baseStr.replace(/[\\/]+$/, "");
  const childClean = childStr.replace(/^[\\/]+/, "");
  return `${baseClean}\\${childClean}`;
}

function buildAutoTemplateFilename(template, date = new Date()) {
  const datePrefix = [
    date.getFullYear(),
    String(date.getMonth() + 1).padStart(2, "0"),
    String(date.getDate()).padStart(2, "0"),
  ].join("-");
  const baseName = String(template?.name || "Template").trim() || "Template";
  return `${datePrefix} ${baseName}`;
}

function buildAutoTemplateDestination(template, selection) {
  const projectPath = String(selection?.project?.path || "").trim();
  if (!projectPath) {
    return { error: "Selected project does not have a path." };
  }
  const discipline = getPrimaryDisciplineLabel();
  const destination = joinPath(projectPath, discipline);
  const newName = buildAutoTemplateFilename(template);
  return { destination, newName };
}

function formatCopyDeliverableLabel(item) {
  const projectLabel =
    item.project?.name || item.project?.nick || item.project?.id || "Untitled Project";
  const deliverableLabel = item.deliverable?.name || "Untitled Deliverable";
  const dueLabel = item.due ? humanDate(item.deliverable?.due) : "";
  const dueSuffix = dueLabel ? ` (Due ${dueLabel})` : "";
  return `${projectLabel} - ${deliverableLabel}${dueSuffix}`;
}

function getCopyDialogDeliverableOptions() {
  const candidates = getAllDeliverables()
    .map(({ project, deliverable }) => {
      const name = String(deliverable?.name || "").trim();
      if (!name) return null;
      const due = parseDueStr(deliverable?.due);
      return { project, deliverable, due };
    })
    .filter(Boolean);

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const weekStart = getWeekStartDate(today);

  const sortByDue = (a, b) => {
    if (!a.due && !b.due) return 0;
    if (!a.due) return 1;
    if (!b.due) return -1;
    return a.due - b.due;
  };

  const withDue = candidates.filter((item) => item.due);
  const upcoming = withDue.filter((item) => item.due >= weekStart);

  if (upcoming.length) return upcoming.sort(sortByDue);
  if (withDue.length) return withDue.sort(sortByDue);
  return candidates;
}

function populateCopyDeliverableSelect() {
  const select = document.getElementById("copy_deliverable_select");
  if (!select) return;
  select.innerHTML = "";
  select.appendChild(
    el("option", { value: "", textContent: "Select deliverable..." })
  );

  copyDialogDeliverables = getCopyDialogDeliverableOptions();
  copyDialogDeliverables.forEach((item, index) => {
    select.appendChild(
      el("option", {
        value: String(index),
        textContent: formatCopyDeliverableLabel(item),
      })
    );
  });

  if (copyDialogDeliverables.length) {
    select.value = "0";
  } else {
    select.value = "";
  }

  updateCopyDialogAutoFields();
}

function getSelectedCopyDeliverable() {
  const select = document.getElementById("copy_deliverable_select");
  if (!select) return null;
  const raw = select.value;
  if (!raw) return null;
  const index = Number(raw);
  if (Number.isNaN(index)) return null;
  return copyDialogDeliverables[index] || null;
}

function updateCopyDialogAutoFields() {
  const destinationInput = document.getElementById("copy_destination");
  const newNameInput = document.getElementById("copy_newName");
  const copyBtn = document.getElementById("btnExecuteCopy");
  copyDialogAutoError = "";

  if (!copyDialogTemplateKey) {
    if (copyBtn) copyBtn.disabled = false;
    return;
  }

  const selection = getSelectedCopyDeliverable();
  if (!selection) {
    if (destinationInput) destinationInput.value = "";
    if (newNameInput) newNameInput.value = "";
    copyDialogAutoError = "Select a deliverable to continue.";
    if (copyBtn) copyBtn.disabled = true;
    return;
  }

  const resolved = buildAutoTemplateDestination(copyDialogTemplate, selection);
  if (resolved.error) {
    if (destinationInput) destinationInput.value = "";
    if (newNameInput) newNameInput.value = "";
    copyDialogAutoError = resolved.error;
    if (copyBtn) copyBtn.disabled = true;
    return;
  }

  if (destinationInput) destinationInput.value = resolved.destination || "";
  if (newNameInput) newNameInput.value = resolved.newName || "";
  if (copyBtn) copyBtn.disabled = false;
}

function setCopyDialogMode(templateKey) {
  const deliverableRow = document.getElementById("copy_deliverable_row");
  const browseBtn = document.getElementById("btnBrowseDestination");
  const newNameInput = document.getElementById("copy_newName");
  const destinationInput = document.getElementById("copy_destination");
  const copyBtn = document.getElementById("btnExecuteCopy");

  const isAuto = Boolean(templateKey);
  if (deliverableRow) deliverableRow.hidden = !isAuto;
  if (browseBtn) browseBtn.disabled = isAuto;
  if (newNameInput) newNameInput.readOnly = isAuto;

  if (destinationInput) {
    destinationInput.value = "";
    destinationInput.placeholder = isAuto
      ? "Auto-set from project and discipline"
      : "Select destination folder...";
  }

  if (newNameInput) {
    newNameInput.value = "";
    if (isAuto) newNameInput.placeholder = "Auto-generated file name";
  }

  if (!isAuto) {
    copyDialogDeliverables = [];
    copyDialogAutoError = "";
    if (copyBtn) copyBtn.disabled = false;
  }
}

async function handleCopyTemplate(templateId) {
  const dlg = document.getElementById('copyTemplateDlg');
  if (!dlg) return;

  const template = templatesDb.templates.find(t => t.id === templateId);
  if (!template) return;

  // Verify template file exists
  const verifyResult = await window.pywebview.api.verify_template_exists(templateId);
  if (verifyResult.status === 'success' && !verifyResult.exists) {
    toast('Template source file not found. The file may have been moved or deleted.');
    return;
  }

  document.getElementById('copy_templateId').value = templateId;
  document.getElementById('copy_templateName').textContent = template.name;
  document.getElementById('copy_destination').value = '';
  document.getElementById('copy_newName').value = '';
  document.getElementById('copy_newName').placeholder = template.name;
  copyDialogTemplate = template;
  copyDialogTemplateKey = getTemplateKey(template);
  setCopyDialogMode(copyDialogTemplateKey);
  if (copyDialogTemplateKey) {
    populateCopyDeliverableSelect();
  }

  dlg.showModal();
}

async function handleBrowseDestination() {
  try {
    const result = await window.pywebview.api.select_folder();
    if (result.status === 'success' && result.path) {
      document.getElementById('copy_destination').value = result.path;
    }
  } catch (e) {
    toast('Error selecting folder.');
  }
}

async function handleExecuteCopy() {
  const templateId = document.getElementById('copy_templateId').value;
  const template = templatesDb.templates.find(t => t.id === templateId);
  if (!template) {
    toast('Template not found.');
    return;
  }

  let destination = document.getElementById('copy_destination').value.trim();
  let newName = document.getElementById('copy_newName').value.trim();
  let context = null;
  let options = null;

  if (copyDialogTemplateKey) {
    const selection = getSelectedCopyDeliverable();
    if (!selection) {
      toast(copyDialogAutoError || 'Please select a deliverable.');
      return;
    }
    const resolved = buildAutoTemplateDestination(template, selection);
    if (resolved.error) {
      toast(resolved.error);
      return;
    }
    destination = resolved.destination;
    newName = resolved.newName;
    context = {
      deliverableName: selection.deliverable?.name || "",
      projectName:
        selection.project?.name ||
        selection.project?.nick ||
        selection.project?.id ||
        "",
    };
    options = {
      createDestination: true,
      templateKey: copyDialogTemplateKey,
    };
  }

  if (!destination) {
    toast('Please select a destination folder.');
    return;
  }

  try {
    const result = await window.pywebview.api.copy_template_to_folder(
      templateId,
      destination,
      newName || null,
      context,
      options
    );

    if (result.status === 'success') {
      toast('Template copied successfully.');
      closeDlg('copyTemplateDlg');

      await window.pywebview.api.open_path(destination);
    } else {
      toast(result.message || 'Failed to copy template.');
    }
  } catch (e) {
    toast('Error copying template.');
  }
}

async function handleRemoveTemplate(templateId, templateName) {
  if (!confirm(`Remove template "${templateName}"?\n\nThis will only remove it from the list, not delete the original file.`)) {
    return;
  }

  try {
    const result = await window.pywebview.api.remove_template(templateId);
    if (result.status === 'success') {
      toast('Template removed.');
      await loadTemplates();
      renderTemplates();
    } else {
      toast(result.message || 'Failed to remove template.');
    }
  } catch (e) {
    toast('Error removing template.');
  }
}

function getTemplateToolContext() {
  let projectPath = "";
  let projectName = "";
  const existingProject = editIndex >= 0 && db[editIndex] ? db[editIndex] : null;
  if (existingProject) {
    projectPath = String(existingProject.path || "").trim();
    projectName = String(existingProject.name || existingProject.nick || existingProject.id || "").trim();
  }

  if (!projectPath) {
    projectPath = document.getElementById("f_path")?.value.trim() || "";
  }
  if (!projectName) {
    projectName = document.getElementById("f_name")?.value.trim() || "";
  }

  return { projectPath, projectName };
}

async function handleTemplateToolSave(templateKey, label) {
  try {
    await loadTemplates();
  } catch (e) {
    console.warn("Failed to refresh templates:", e);
  }

  const template = getTemplateByKey(templateKey);
  if (!template) {
    toast(`Template "${label}" not found.`);
    return;
  }

  const { projectPath, projectName } = getTemplateToolContext();
  const defaultName = String(label || template.name || "Template").trim() || "Template";
  let selection = null;
  try {
    selection = await window.pywebview.api.select_template_save_location(
      projectPath || null,
      defaultName,
      template.fileType
    );
  } catch (e) {
    toast("Error selecting save location.");
    return;
  }

  if (!selection || selection.status === "cancelled") {
    return;
  }
  if (selection.status !== "success" || !selection.path) {
    toast(selection.message || "Failed to select save location.");
    return;
  }

  const context = {};
  if (projectName) context.projectName = projectName;
  const options = { templateKey };

  try {
    const result = await window.pywebview.api.copy_template_to_path(
      template.id,
      selection.path,
      context,
      options
    );
    if (result.status === "success") {
      toast("Template saved.");
    } else {
      toast(result.message || "Failed to save template.");
    }
  } catch (e) {
    toast("Error saving template.");
  }
}

// ===================== TIMESHEET RENDERING =====================

function renderTimesheets() {
  const weekKey = formatWeekKey(currentTimesheetWeek);
  const entries = getWeekEntries(weekKey);

  // Update week display
  const weekDisplay = document.getElementById("weekDisplay");
  if (weekDisplay) {
    weekDisplay.textContent = formatWeekDisplay(currentTimesheetWeek);
  }

  renderTimesheetEntries(entries);
  renderTimesheetSuggestions();
  updateTimesheetTotals(entries);
  renderExpenses(); // Also render expense sheet when timesheets render
}

function renderTimesheetEntries(entries) {
  const tbody = document.getElementById("timesheetBody");
  if (!tbody) return;
  tbody.innerHTML = "";

  const emptyState = document.getElementById("timesheetEmptyState");
  if (!entries.length) {
    if (emptyState) emptyState.style.display = "block";
    return;
  }
  if (emptyState) emptyState.style.display = "none";

  ensureTimesheetEntryIds(entries);
  entries.forEach((entry, index) => {
    const row = createTimesheetRow(entry, index);
    tbody.appendChild(row);
  });
}

function createTimesheetRow(entry, index) {
  const row = el("tr", { className: "ts-entry-row" });
  row.dataset.entryId = entry.id || "";

  const dragCell = el("td", { className: "ts-drag-col" });
  const dragHandle = el("button", {
    className: "ts-drag-handle",
    type: "button",
    title: "Drag to reorder",
    "aria-label": "Drag to reorder",
    textContent: "|||",
  });
  dragCell.appendChild(dragHandle);
  row.appendChild(dragCell);

  // Project ID
  const projectIdCell = el("td");
  const projectIdInput = el("input", {
    value: entry.projectId || "",
    placeholder: "--",
  });
  projectIdInput.oninput = (e) => {
    entry.projectId = e.target.value;
    saveTimesheets();
  };
  projectIdCell.appendChild(projectIdInput);
  row.appendChild(projectIdCell);

  // Task Number (editable)
  const taskCell = el("td");
  const taskInput = el("input", {
    value: entry.taskNumber || "",
    placeholder: "--",
  });
  taskInput.oninput = (e) => {
    entry.taskNumber = e.target.value;
    saveTimesheets();
  };
  taskCell.appendChild(taskInput);
  row.appendChild(taskCell);

  // Project Name
  const nameCell = el("td", { className: "ts-name-col" });
  const nameInput = el("input", {
    value: entry.projectName || "",
    placeholder: "--",
  });
  nameInput.oninput = (e) => {
    entry.projectName = e.target.value;
    saveTimesheets();
  };
  nameCell.appendChild(nameInput);
  row.appendChild(nameCell);

  // Function (auto-filled, editable)
  const funcCell = el("td");
  const funcInput = el("input", {
    value: entry.function || getDisciplineFunction(),
    style: "text-align: center",
  });
  funcInput.oninput = (e) => {
    entry.function = e.target.value.toUpperCase();
    saveTimesheets();
  };
  funcCell.appendChild(funcInput);
  row.appendChild(funcCell);

  // PM Initials (editable)
  const pmCell = el("td");
  const pmInput = el("input", {
    value: entry.pmInitials || "",
    placeholder: "--",
    style: "text-align: center",
  });
  pmInput.oninput = (e) => {
    entry.pmInitials = e.target.value.toUpperCase();
    saveTimesheets();
  };
  pmCell.appendChild(pmInput);
  row.appendChild(pmCell);

  // Service Description (auto-filled from deliverable, editable)
  const descCell = el("td");
  const descInput = el("input", {
    value: entry.serviceDescription || entry.deliverableName || "",
    placeholder: "Description",
  });
  descInput.oninput = (e) => {
    entry.serviceDescription = e.target.value;
    saveTimesheets();
  };
  descCell.appendChild(descInput);
  row.appendChild(descCell);

  // Hour cells for each day (Mon-Sun)
  const dayOrder = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"];
  dayOrder.forEach((day) => {
    const cell = el("td", { className: "ts-hour-cell" });
    const hours = entry.hours?.[day] || 0;
    cell.appendChild(createHourDragBar(entry, day, hours, row));
    row.appendChild(cell);
  });

  // Row Total
  const rowTotal = calculateRowTotal(entry);
  row.appendChild(
    el("td", {
      className: "ts-row-total",
      textContent: rowTotal > 0 ? rowTotal.toFixed(1) : "",
    })
  );

  // Mileage
  const mileageCell = el("td");
  const mileageInput = el("input", {
    type: "number",
    value: entry.mileage || "",
    placeholder: "0",
    min: "0",
    style: "text-align: center",
  });
  mileageInput.oninput = (e) => {
    entry.mileage = parseFloat(e.target.value) || 0;
    saveTimesheets();
    updateTimesheetTotals(getWeekEntries(formatWeekKey(currentTimesheetWeek)));
  };
  mileageCell.appendChild(mileageInput);
  row.appendChild(mileageCell);

  // Remove button
  const actionsCell = el("td");
  const removeBtn = el("button", {
    className: "ts-remove-btn",
    innerHTML: "&times;",
    title: "Remove entry",
  });
  removeBtn.onclick = () => removeTimesheetEntry(index);
  actionsCell.appendChild(removeBtn);
  row.appendChild(actionsCell);

  enableTimesheetRowDrag(row, entry.id);
  return row;
}

function enableTimesheetRowDrag(row, entryId) {
  const handle = row.querySelector(".ts-drag-handle");
  if (!handle) return;
  handle.draggable = true;

  handle.addEventListener("dragstart", (e) => {
    timesheetDragEntryId = entryId;
    row.classList.add("ts-dragging");
    if (e.dataTransfer) {
      e.dataTransfer.effectAllowed = "move";
      e.dataTransfer.setData("text/plain", entryId);
    }
  });

  handle.addEventListener("dragend", () => {
    timesheetDragEntryId = null;
    row.classList.remove("ts-dragging");
    clearTimesheetDragStyles();
  });

  row.addEventListener("dragover", (e) => {
    if (!timesheetDragEntryId) return;
    e.preventDefault();
    const rect = row.getBoundingClientRect();
    const before = e.clientY < rect.top + rect.height / 2;
    row.classList.toggle("ts-drop-before", before);
    row.classList.toggle("ts-drop-after", !before);
    row.dataset.dropPosition = before ? "before" : "after";
  });

  row.addEventListener("dragleave", () => {
    row.classList.remove("ts-drop-before", "ts-drop-after");
    delete row.dataset.dropPosition;
  });

  row.addEventListener("drop", (e) => {
    if (!timesheetDragEntryId) return;
    e.preventDefault();
    const before = row.dataset.dropPosition !== "after";
    const weekKey = formatWeekKey(currentTimesheetWeek);
    const entries = getWeekEntries(weekKey);
    const fromIndex = entries.findIndex(
      (entry) => entry.id === timesheetDragEntryId
    );
    const toIndex = entries.findIndex((entry) => entry.id === entryId);
    if (fromIndex < 0 || toIndex < 0 || fromIndex === toIndex) {
      clearTimesheetDragStyles();
      timesheetDragEntryId = null;
      return;
    }
    let targetIndex = before ? toIndex : toIndex + 1;
    if (fromIndex < targetIndex) targetIndex -= 1;
    moveTimesheetEntry(entries, fromIndex, targetIndex);
    setWeekEntries(weekKey, entries);
    saveTimesheets();
    renderTimesheets();
    clearTimesheetDragStyles();
    timesheetDragEntryId = null;
  });

}

function createHourDragBar(entry, day, hours, row) {
  const container = el("div", { className: "ts-hour-bar-container" });

  const fill = el("div", {
    className: "ts-hour-bar-fill",
    style: `width: ${(hours / MAX_HOURS_PER_DAY) * 100}%`,
  });

  const valueDisplay = el("div", {
    className: "ts-hour-value",
    textContent: hours > 0 ? hours.toFixed(1) : "",
  });

  container.append(fill, valueDisplay);

  // Drag interaction
  let isDragging = false;

  const updateHours = (e) => {
    const rect = container.getBoundingClientRect();
    const x = Math.max(0, Math.min(e.clientX - rect.left, rect.width));
    const percentage = x / rect.width;
    const newHours = Math.round(percentage * MAX_HOURS_PER_DAY * 2) / 2; // Round to 0.5

    if (!entry.hours) entry.hours = {};
    entry.hours[day] = Math.min(newHours, MAX_HOURS_PER_DAY);

    fill.style.width = `${(entry.hours[day] / MAX_HOURS_PER_DAY) * 100}%`;
    valueDisplay.textContent =
      entry.hours[day] > 0 ? entry.hours[day].toFixed(1) : "";

    // Update row total
    const rowTotalCell = row?.querySelector(".ts-row-total");
    if (rowTotalCell) {
      const total = calculateRowTotal(entry);
      rowTotalCell.textContent = total > 0 ? total.toFixed(1) : "";
    }

    updateTimesheetTotals(getWeekEntries(formatWeekKey(currentTimesheetWeek)));
  };

  container.onmousedown = (e) => {
    isDragging = true;
    updateHours(e);
    e.preventDefault();
  };

  const onMouseMove = (e) => {
    if (isDragging) updateHours(e);
  };

  const onMouseUp = () => {
    if (isDragging) {
      isDragging = false;
      saveTimesheets();
    }
  };

  document.addEventListener("mousemove", onMouseMove);
  document.addEventListener("mouseup", onMouseUp);

  // Double-click to manually enter
  container.ondblclick = () => {
    const value = prompt("Enter hours:", entry.hours?.[day] || 0);
    if (value !== null) {
      const newHours = Math.min(parseFloat(value) || 0, MAX_HOURS_PER_DAY);
      if (!entry.hours) entry.hours = {};
      entry.hours[day] = newHours;
      fill.style.width = `${(newHours / MAX_HOURS_PER_DAY) * 100}%`;
      valueDisplay.textContent = newHours > 0 ? newHours.toFixed(1) : "";
      saveTimesheets();
      updateTimesheetTotals(getWeekEntries(formatWeekKey(currentTimesheetWeek)));
      // Update row total
      const rowTotalCell = row?.querySelector(".ts-row-total");
      if (rowTotalCell) {
        const total = calculateRowTotal(entry);
        rowTotalCell.textContent = total > 0 ? total.toFixed(1) : "";
      }
    }
  };

  return container;
}

function calculateRowTotal(entry) {
  if (!entry.hours) return 0;
  return Object.values(entry.hours).reduce((sum, h) => sum + (h || 0), 0);
}

function updateTimesheetTotals(entries) {
  const dayOrder = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"];
  const totals = { week: 0, mileage: 0 };

  dayOrder.forEach((day) => (totals[day] = 0));

  entries.forEach((entry) => {
    dayOrder.forEach((day) => {
      totals[day] += entry.hours?.[day] || 0;
    });
    totals.mileage += entry.mileage || 0;
  });

  totals.week = dayOrder.reduce((sum, day) => sum + totals[day], 0);

  // Update footer cells
  dayOrder.forEach((day) => {
    const cell = document.getElementById(
      `total${day.charAt(0).toUpperCase() + day.slice(1)}`
    );
    if (cell) cell.textContent = totals[day] > 0 ? totals[day].toFixed(1) : "0";
  });

  const weekCell = document.getElementById("totalWeek");
  if (weekCell)
    weekCell.textContent = totals.week > 0 ? totals.week.toFixed(1) : "0";

  const mileageCell = document.getElementById("totalMileage");
  if (mileageCell) mileageCell.textContent = totals.mileage || "0";

  const weeklyTotal = document.getElementById("weeklyTotalHours");
  if (weeklyTotal)
    weeklyTotal.textContent = totals.week > 0 ? totals.week.toFixed(1) : "0.0";
}

function renderTimesheetSuggestions() {
  const container = document.getElementById("timesheetSuggestions");
  if (!container) return;
  container.innerHTML = "";

  const suggestions = getProjectsForWeek(currentTimesheetWeek);
  const weekKey = formatWeekKey(currentTimesheetWeek);
  const entries = getWeekEntries(weekKey);

  const weeklyCard = createWeeklyMeetingCard();
  if (weeklyCard) container.appendChild(weeklyCard);

  if (!suggestions.length) {
    container.appendChild(
      el("p", { className: "muted", textContent: "No projects due this week or next." })
    );
    return;
  }

  const summaryEntriesByProject = new Map();
  entries.forEach((entry) => {
    if (!isProjectSummaryEntry(entry)) return;
    const key = getTimesheetEntryMatchKey(entry);
    if (!key) return;
    const list = summaryEntriesByProject.get(key) || [];
    list.push(entry);
    summaryEntriesByProject.set(key, list);
  });

  const grouped = new Map();
  let fallbackIndex = 0;

  suggestions.forEach(({ project, deliverable, dueDate }) => {
    let key = getTimesheetProjectMatchKey(project);
    if (!key) {
      key = getTimesheetProjectMatchKey(project, fallbackIndex);
      fallbackIndex += 1;
    }
    const group = grouped.get(key) || {
      project,
      deliverables: [],
      earliestDueDate: null,
      earliestDue: "",
    };
    group.deliverables.push(deliverable);
    if (!group.earliestDueDate || (dueDate && dueDate < group.earliestDueDate)) {
      group.earliestDueDate = dueDate || group.earliestDueDate;
      group.earliestDue = deliverable?.due || group.earliestDue;
    }
    grouped.set(key, group);
  });

  const thisWeekEnd = new Date(currentTimesheetWeek);
  thisWeekEnd.setDate(thisWeekEnd.getDate() + 6);

  const groupedList = Array.from(grouped.values()).sort((a, b) => {
    const aDate = a.earliestDueDate;
    const bDate = b.earliestDueDate;
    const aThisWeek = aDate && aDate <= thisWeekEnd;
    const bThisWeek = bDate && bDate <= thisWeekEnd;
    if (aThisWeek && !bThisWeek) return -1;
    if (!aThisWeek && bThisWeek) return 1;
    if (!aDate && !bDate) return 0;
    if (!aDate) return 1;
    if (!bDate) return -1;
    return aDate - bDate;
  });

  groupedList.forEach((group) => {
    const deliverableNames = normalizeDeliverableNames(group.deliverables);
    const matchKey = getTimesheetProjectMatchKey(group.project);
    const summaryEntries = matchKey
      ? summaryEntriesByProject.get(matchKey) || []
      : [];
    const entryStatus = getSummaryEntryStatus(
      summaryEntries,
      deliverableNames
    );
    const card = createSuggestionCard(
      group.project,
      deliverableNames,
      group.earliestDue,
      entryStatus
    );
    container.appendChild(card);
  });
}

function createWeeklyMeetingCard() {
  const card = el("div", { className: "ts-suggestion-card" });
  card.onclick = () => addWeeklyMeetingToTimesheet();

  const header = el("div", { className: "ts-suggestion-header" });
  header.appendChild(
    el("span", { className: "ts-suggestion-id", textContent: "--" })
  );
  header.appendChild(
    el("span", {
      className: "ts-suggestion-due ok",
      textContent: "Weekly",
    })
  );
  card.appendChild(header);

  card.appendChild(
    el("div", {
      className: "ts-suggestion-name",
      textContent: WEEKLY_MEETING_NAME,
    })
  );

  card.appendChild(
    el("div", {
      className: "ts-suggestion-deliverable",
      textContent: WEEKLY_MEETING_NAME,
    })
  );

  return card;
}

function createSuggestionCard(
  project,
  deliverableNames,
  dueDate,
  entryStatus = {}
) {
  const {
    count = 0,
    needsUpdate = false,
  } = entryStatus;
  const showAddedStyle = count > 0 && !needsUpdate;
  const card = el("div", {
    className: `ts-suggestion-card ${showAddedStyle ? "added" : ""}`,
  });

  card.onclick = () => {
    addProjectDeliverablesToTimesheet(project, deliverableNames);
  };

  const header = el("div", { className: "ts-suggestion-header" });
  header.appendChild(
    el("span", { className: "ts-suggestion-id", textContent: project.id || "--" })
  );

  const ds = dueState(dueDate);
  header.appendChild(
    el("span", {
      className: `ts-suggestion-due ${ds}`,
      textContent: humanDate(dueDate),
    })
  );
  card.appendChild(header);

  card.appendChild(
    el("div", {
      className: "ts-suggestion-name",
      textContent: project.nick || project.name || "Unnamed Project",
    })
  );

  card.appendChild(
    (() => {
      const deliverableEl = el("div", { className: "ts-suggestion-deliverable" });
      if (!deliverableNames.length) {
        deliverableEl.textContent = "No deliverables listed";
        return deliverableEl;
      }
      deliverableNames.forEach((name) => {
        deliverableEl.appendChild(el("div", { textContent: name }));
      });
      return deliverableEl;
    })()
  );

  if (needsUpdate) {
    card.appendChild(
      el("div", {
        className: "ts-suggestion-added",
        textContent: "Click to append new deliverables",
      })
    );
  } else if (count > 0) {
    card.appendChild(
      el("div", {
        className: "ts-suggestion-added",
        textContent: "Added to timesheet",
      })
    );
  }

  return card;
}

// ===================== TIMESHEET ACTIONS =====================

function buildProjectSummaryEntry({
  projectId,
  projectName,
  pmInitials,
  functionName,
  taskNumber,
  description,
  isWfh,
}) {
  const baseDescription = description || "";
  const serviceDescription = isWfh
    ? appendWfhSuffix(baseDescription)
    : baseDescription;

  return {
    id: createTimesheetEntryId(),
    projectId,
    projectName,
    deliverableId: TIMESHEET_SUMMARY_DELIVERABLE_ID,
    deliverableName: serviceDescription,
    pmInitials,
    serviceDescription,
    function: functionName,
    hours: { sun: 0, mon: 0, tue: 0, wed: 0, thu: 0, fri: 0, sat: 0 },
    mileage: 0,
    taskNumber,
    isProjectSummary: true,
  };
}

function addWeeklyMeetingToTimesheet() {
  const weekKey = formatWeekKey(currentTimesheetWeek);
  const entries = getWeekEntries(weekKey);
  const hours = {
    sun: 0,
    mon: WEEKLY_MEETING_DEFAULT_HOURS,
    tue: 0,
    wed: 0,
    thu: 0,
    fri: 0,
    sat: 0,
  };
  const taskNumber = getTimesheetTaskNumberForDeliverable(WEEKLY_MEETING_NAME);

  entries.push({
    id: createTimesheetEntryId(),
    projectId: "",
    projectName: WEEKLY_MEETING_NAME,
    deliverableId: "",
    deliverableName: WEEKLY_MEETING_NAME,
    pmInitials: getDefaultPmInitials(),
    serviceDescription: WEEKLY_MEETING_NAME,
    function: getDisciplineFunction(),
    hours,
    mileage: 0,
    taskNumber,
  });

  setWeekEntries(weekKey, entries);
  saveTimesheets();
  renderTimesheets();
  toast("Added workload meeting entry.");
}

function addProjectDeliverablesToTimesheet(project, deliverables) {
  const weekKey = formatWeekKey(currentTimesheetWeek);
  const entries = getWeekEntries(weekKey);
  const deliverableNames = normalizeDeliverableNames(deliverables);
  const projectKey = getTimesheetProjectMatchKey(project);

  if (!projectKey) {
    toast("Project is missing an ID or name.");
    return;
  }

  const baseDescription = deliverableNames.join(", ");

  if (!baseDescription) {
    toast("No deliverables due this week.");
    return;
  }

  const summaryEntries = entries.filter(
    (entry) =>
      isProjectSummaryEntry(entry) &&
      getTimesheetEntryMatchKey(entry) === projectKey
  );

  const projectId = project.id || "";
  const projectName = project.nick || project.name || "";

  const updateEntryDescription = (entry) => {
    const currentDesc = getTimesheetEntryDescription(entry);
    const isWfh = hasWfhSuffix(currentDesc);
    const baseDesc = stripWfhSuffix(currentDesc);
    const merged = mergeDeliverableDescription(baseDesc, deliverableNames);
    if (merged !== baseDesc) {
      const nextDesc = isWfh ? appendWfhSuffix(merged) : merged;
      entry.serviceDescription = nextDesc;
      entry.deliverableName = nextDesc;
    }
  };

  if (summaryEntries.length) summaryEntries.forEach(updateEntryDescription);

  const referenceEntry = summaryEntries[0];

  entries.push(
    buildProjectSummaryEntry({
      projectId: referenceEntry?.projectId || projectId,
      projectName: referenceEntry?.projectName || projectName,
      pmInitials: referenceEntry?.pmInitials || getDefaultPmInitials(),
      functionName: referenceEntry?.function || getDisciplineFunction(),
      taskNumber: getTimesheetTaskNumberForDeliverable(baseDescription),
      description: baseDescription,
      isWfh: false,
    })
  );

  setWeekEntries(weekKey, entries);
  saveTimesheets();
  renderTimesheets();

  toast("Added timesheet entry.");
}

function addProjectToTimesheet(project, deliverable) {
  const weekKey = formatWeekKey(currentTimesheetWeek);
  const entries = getWeekEntries(weekKey);
  const projectId = project.id || "";
  const deliverableId = deliverable?.id || "";
  const existingEntries = entries.filter(
    (e) => e.projectId === projectId && e.deliverableId === deliverableId
  );

  const referenceEntry = existingEntries[0];
  const baseDescription = referenceEntry
    ? stripWfhSuffix(
      getTimesheetEntryDescription(referenceEntry, deliverable?.name || "")
    )
    : deliverable?.name || "";
  const serviceDescription = baseDescription;
  const projectName =
    referenceEntry?.projectName || project.nick || project.name || "";
  const deliverableName =
    deliverable?.name || referenceEntry?.deliverableName || "";
  const pmInitials = referenceEntry?.pmInitials || getDefaultPmInitials();
  const functionName = referenceEntry?.function || getDisciplineFunction();
  const taskNumber = getTimesheetTaskNumberForDeliverable(
    deliverableName || baseDescription
  );

  const entry = {
    id: createTimesheetEntryId(),
    projectId,
    projectName,
    deliverableId,
    deliverableName,
    pmInitials,
    serviceDescription,
    function: functionName,
    hours: { sun: 0, mon: 0, tue: 0, wed: 0, thu: 0, fri: 0, sat: 0 },
    mileage: 0,
    taskNumber,
  };

  entries.push(entry);
  setWeekEntries(weekKey, entries);
  toast("Added timesheet entry.");
  saveTimesheets();
  renderTimesheets();
}

function addManualTimesheetEntry() {
  const weekKey = formatWeekKey(currentTimesheetWeek);
  const entries = getWeekEntries(weekKey);

  const entry = {
    id: createTimesheetEntryId(),
    projectId: "",
    projectName: "",
    deliverableId: "",
    deliverableName: "",
    pmInitials: getDefaultPmInitials(),
    serviceDescription: "",
    function: getDisciplineFunction(),
    hours: { sun: 0, mon: 0, tue: 0, wed: 0, thu: 0, fri: 0, sat: 0 },
    mileage: 0,
    taskNumber: "0.00",
  };

  entries.push(entry);
  setWeekEntries(weekKey, entries);
  toast("Added manual timesheet entry.");
  saveTimesheets();
  renderTimesheets();
}

function removeTimesheetEntry(index) {
  const weekKey = formatWeekKey(currentTimesheetWeek);
  const entries = getWeekEntries(weekKey);
  entries.splice(index, 1);
  setWeekEntries(weekKey, entries);
  saveTimesheets();
  renderTimesheets();
}

function openAddTimesheetProjectDialog() {
  const dlg = document.getElementById("timesheetProjectDlg");
  const list = document.getElementById("timesheetProjectList");
  const search = document.getElementById("timesheetProjectSearch");

  if (!dlg || !list) return;

  const renderProjects = (query = "") => {
    list.innerHTML = "";
    const q = query.toLowerCase();

    db.filter((p) => {
      const name = (p.name || "").toLowerCase();
      const id = (p.id || "").toLowerCase();
      const nick = (p.nick || "").toLowerCase();
      return !q || name.includes(q) || id.includes(q) || nick.includes(q);
    }).forEach((project) => {
      const deliverables = getProjectDeliverables(project);
      deliverables.forEach((deliverable) => {
        const option = el("div", { className: "ts-project-option" });
        option.innerHTML = `
          <div><strong>${project.id || "--"}</strong> - ${project.nick || project.name || "Unnamed"
          }</div>
          <div class="muted tiny">${deliverable.name || "Deliverable"} - Due: ${humanDate(deliverable.due) || "No date"
          }</div>
        `;
        option.onclick = () => {
          addProjectToTimesheet(project, deliverable);
          closeDlg("timesheetProjectDlg");
        };
        list.appendChild(option);
      });
    });

    if (list.children.length === 0) {
      list.innerHTML = '<p class="muted" style="padding: 1rem; text-align: center;">No projects found</p>';
    }
  };

  renderProjects();
  if (search) {
    search.value = "";
    search.oninput = () => renderProjects(search.value);
  }

  showDialog(dlg);
}

function navigateTimesheetWeek(direction) {
  const newDate = new Date(currentTimesheetWeek);
  newDate.setDate(newDate.getDate() + direction * 7);
  currentTimesheetWeek = newDate;
  renderTimesheets();
}

function goToCurrentWeek() {
  currentTimesheetWeek = getWeekStartDate(new Date());
  renderTimesheets();
}

function calculateWeekTotals(entries) {
  const dayOrder = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"];
  const totals = { mileage: 0 };
  dayOrder.forEach((day) => (totals[day] = 0));

  entries.forEach((entry) => {
    dayOrder.forEach((day) => (totals[day] += entry.hours?.[day] || 0));
    totals.mileage += entry.mileage || 0;
  });

  totals.week = dayOrder.reduce((sum, day) => sum + totals[day], 0);
  return totals;
}

async function exportTimesheetToExcel() {
  const weekKey = formatWeekKey(currentTimesheetWeek);
  const entries = getWeekEntries(weekKey);
  const expenseProjects = getWeekExpenses(weekKey);

  if (!entries.length && !expenseProjects.length) {
    toast("No entries to export");
    return;
  }

  try {
    // Export timesheet and expenses together
    const result = await window.pywebview.api.export_timesheet_excel({
      weekKey,
      weekDisplay: formatWeekDisplay(currentTimesheetWeek),
      userName: userSettings.userName || "Employee",
      entries,
      totals: calculateWeekTotals(entries),
      // Include expense data for the expense sheet tab
      expenses: {
        projects: expenseProjects,
        mileageRate: MILEAGE_RATE
      }
    });

    if (result.status === "cancelled") {
      toast("Export cancelled.");
      return;
    }
    if (result.status === "success") {
      const hasTimesheet = entries.length > 0;
      const hasExpenses = expenseProjects.length > 0;
      const message = hasTimesheet && hasExpenses
        ? "Timesheet and project expense sheet exported successfully"
        : hasExpenses
          ? "Project expense sheet exported successfully"
          : "Timesheet exported successfully";
      toast(message);
    } else {
      throw new Error(result.message);
    }
  } catch (e) {
    console.error("Export failed:", e);
    toast("Failed to export: " + e.message);
  }
}

// ===================== EXPENSE SHEET FUNCTIONS =====================

const MILEAGE_RATE = 0.70; // Fixed rate per mile

function createExpenseEntryId() {
  return "exp_" + Math.random().toString(36).substr(2, 9);
}

function createExpenseProjectId() {
  return "prj_" + Math.random().toString(36).substr(2, 9);
}

function createExpenseImageId() {
  return "img_" + Math.random().toString(36).substr(2, 9);
}

function getWeekExpenses(weekKey) {
  return timesheetDb.expenses?.[weekKey]?.projects || [];
}

function setWeekExpenses(weekKey, projects) {
  if (!timesheetDb.expenses) {
    timesheetDb.expenses = {};
  }
  if (!timesheetDb.expenses[weekKey]) {
    timesheetDb.expenses[weekKey] = { projects: [] };
  }
  timesheetDb.expenses[weekKey].projects = projects;
}

function renderExpenses() {
  const weekKey = formatWeekKey(currentTimesheetWeek);
  const projects = getWeekExpenses(weekKey);

  const container = document.getElementById("expenseSheetContainer");
  const emptyState = document.getElementById("expenseSheetEmptyState");
  const totalsBar = document.getElementById("expenseTotalsBar");

  if (!container) return;
  container.innerHTML = "";

  if (!projects.length) {
    if (emptyState) emptyState.style.display = "block";
    if (totalsBar) totalsBar.style.display = "none";
    return;
  }

  if (emptyState) emptyState.style.display = "none";
  if (totalsBar) totalsBar.style.display = "flex";

  projects.forEach((project, projectIndex) => {
    const section = createExpenseProjectSection(project, projectIndex);
    container.appendChild(section);
  });

  updateExpenseTotals(projects);
}

function createExpenseProjectSection(project, projectIndex) {
  const section = el("div", { className: "expense-project" });
  section.dataset.projectIndex = projectIndex;

  // Header
  const header = el("div", { className: "expense-project-header" });

  const info = el("div", { className: "expense-project-info" });
  info.appendChild(el("div", { className: "expense-project-name", textContent: project.projectName || "Unnamed Project" }));
  info.appendChild(el("div", { className: "expense-project-id", textContent: `Job #: ${project.projectId || "--"}` }));
  header.appendChild(info);

  const actions = el("div", { className: "expense-project-actions" });

  const addEntryBtn = el("button", {
    className: "btn tiny",
    textContent: "+ Add Entry",
    onclick: () => openAddExpenseEntryDialog(projectIndex)
  });
  actions.appendChild(addEntryBtn);

  const addImageBtn = el("button", {
    className: "btn tiny",
    textContent: "📷 Add Images",
    onclick: () => addExpenseImages(projectIndex)
  });
  actions.appendChild(addImageBtn);

  const removeProjectBtn = el("button", {
    className: "btn tiny ghost text-danger",
    textContent: "Remove",
    onclick: () => removeExpenseProject(projectIndex)
  });
  actions.appendChild(removeProjectBtn);

  header.appendChild(actions);
  section.appendChild(header);

  // Expense Table
  const tableContainer = el("div", { className: "expense-table-container" });
  const table = el("table", { className: "expense-table" });

  // Table Header
  const thead = el("thead");
  const headerRow = el("tr");
  headerRow.appendChild(el("th", { className: "date-col", textContent: "Date" }));
  headerRow.appendChild(el("th", { textContent: "Description" }));
  headerRow.appendChild(el("th", { className: "mileage-col", textContent: "Mileage" }));
  headerRow.appendChild(el("th", { className: "amount-col", textContent: "Expense" }));
  headerRow.appendChild(el("th", { className: "actions-col", textContent: "" }));
  thead.appendChild(headerRow);
  table.appendChild(thead);

  // Table Body
  const tbody = el("tbody");
  (project.entries || []).forEach((entry, entryIndex) => {
    const row = createExpenseRow(entry, projectIndex, entryIndex);
    tbody.appendChild(row);
  });

  // Subtotals row
  const subtotalsRow = el("tr", { className: "subtotals-row" });
  const projectTotals = calculateProjectExpenseTotals(project);
  subtotalsRow.appendChild(el("td", { colSpan: 2, textContent: "Subtotal" }));
  subtotalsRow.appendChild(el("td", { className: "mileage-col", textContent: projectTotals.mileage.toString() }));
  subtotalsRow.appendChild(el("td", { className: "amount-col", textContent: "$" + projectTotals.expense.toFixed(2) }));
  subtotalsRow.appendChild(el("td", { className: "actions-col" }));
  tbody.appendChild(subtotalsRow);

  table.appendChild(tbody);
  tableContainer.appendChild(table);
  section.appendChild(tableContainer);

  // Images Section
  const imagesSection = el("div", { className: "expense-images-section" });

  const imagesHeader = el("div", { className: "expense-images-header" });
  imagesHeader.appendChild(el("span", { className: "expense-images-title", textContent: "Receipts & Attachments" }));
  imagesSection.appendChild(imagesHeader);

  const imagesGrid = el("div", { className: "expense-images-grid" });

  if (project.images && project.images.length > 0) {
    project.images.forEach((image, imageIndex) => {
      const thumb = createExpenseImageThumb(image, projectIndex, imageIndex);
      imagesGrid.appendChild(thumb);
    });
  } else {
    imagesGrid.appendChild(el("span", { className: "expense-images-empty", textContent: "No receipts attached" }));
  }

  imagesSection.appendChild(imagesGrid);
  section.appendChild(imagesSection);

  return section;
}

function createExpenseRow(entry, projectIndex, entryIndex) {
  const row = el("tr");
  row.dataset.entryIndex = entryIndex;

  // Date
  const dateCell = el("td", { className: "date-col" });
  const dateInput = el("input", {
    type: "date",
    value: entry.date || "",
  });
  dateInput.oninput = (e) => {
    entry.date = e.target.value;
    saveTimesheets();
  };
  dateCell.appendChild(dateInput);
  row.appendChild(dateCell);

  // Description
  const descCell = el("td");
  const descInput = el("input", {
    value: entry.description || "",
    placeholder: "Description"
  });
  descInput.oninput = (e) => {
    entry.description = e.target.value;
    saveTimesheets();
  };
  descCell.appendChild(descInput);
  row.appendChild(descCell);

  // Mileage
  const mileageCell = el("td", { className: "mileage-col" });
  const mileageInput = el("input", {
    type: "number",
    min: "0",
    value: entry.mileage || "",
    placeholder: "0"
  });
  mileageInput.oninput = (e) => {
    entry.mileage = parseFloat(e.target.value) || 0;
    saveTimesheets();
    updateAllExpenseTotals(projectIndex);
  };
  mileageCell.appendChild(mileageInput);
  row.appendChild(mileageCell);

  // Expense Amount
  const amountCell = el("td", { className: "amount-col" });
  const amountInput = el("input", {
    type: "number",
    step: "0.01",
    min: "0",
    value: entry.expense || "",
    placeholder: "0.00"
  });
  amountInput.oninput = (e) => {
    entry.expense = parseFloat(e.target.value) || 0;
    saveTimesheets();
    updateAllExpenseTotals(projectIndex);
  };
  amountCell.appendChild(amountInput);
  row.appendChild(amountCell);

  // Remove button
  const actionsCell = el("td", { className: "actions-col" });
  const removeBtn = el("button", {
    className: "expense-remove-btn",
    innerHTML: "&times;",
    title: "Remove entry"
  });
  removeBtn.onclick = () => removeExpenseEntry(projectIndex, entryIndex);
  actionsCell.appendChild(removeBtn);
  row.appendChild(actionsCell);

  return row;
}

function createExpenseImageThumb(image, projectIndex, imageIndex) {
  const thumb = el("div", { className: "expense-image-thumb" });

  const ext = (image.filename || "").toLowerCase().split(".").pop();

  if (ext === "pdf") {
    thumb.appendChild(el("div", { className: "pdf-icon", textContent: "📄" }));
  } else {
    // For images, we can't actually load local file:// URLs due to security
    // So we'll show a placeholder or use the filename
    const imgEl = el("div", {
      className: "pdf-icon",
      textContent: "🖼️",
      title: image.filename || "Image"
    });
    thumb.appendChild(imgEl);
  }

  const removeBtn = el("button", {
    className: "remove-image-btn",
    innerHTML: "&times;",
    title: "Remove image"
  });
  removeBtn.onclick = (e) => {
    e.stopPropagation();
    removeExpenseImage(projectIndex, imageIndex);
  };
  thumb.appendChild(removeBtn);

  return thumb;
}

function calculateProjectExpenseTotals(project) {
  let mileage = 0;
  let expense = 0;

  (project.entries || []).forEach((entry) => {
    mileage += entry.mileage || 0;
    expense += entry.expense || 0;
  });

  return { mileage, expense };
}

function updateExpenseTotals(projects) {
  let totalMileage = 0;
  let totalExpense = 0;

  projects.forEach((project) => {
    const totals = calculateProjectExpenseTotals(project);
    totalMileage += totals.mileage;
    totalExpense += totals.expense;
  });

  const mileageReimbursement = totalMileage * MILEAGE_RATE;
  const grandTotal = mileageReimbursement + totalExpense;

  const mileageEl = document.getElementById("expenseTotalMileage");
  const reimbursementEl = document.getElementById("expenseMileageReimbursement");
  const expenseEl = document.getElementById("expenseTotalAmount");
  const grandTotalEl = document.getElementById("expenseGrandTotal");

  if (mileageEl) mileageEl.textContent = totalMileage.toString();
  if (reimbursementEl) reimbursementEl.textContent = "$" + mileageReimbursement.toFixed(2);
  if (expenseEl) expenseEl.textContent = "$" + totalExpense.toFixed(2);
  if (grandTotalEl) grandTotalEl.textContent = "$" + grandTotal.toFixed(2);
}

function updateAllExpenseTotals(projectIndex) {
  const weekKey = formatWeekKey(currentTimesheetWeek);
  const projects = getWeekExpenses(weekKey);

  // Update subtotals row for the specific project
  const projectSection = document.querySelector(`.expense-project[data-project-index="${projectIndex}"]`);
  if (projectSection && projects[projectIndex]) {
    const subtotalsRow = projectSection.querySelector(".subtotals-row");
    if (subtotalsRow) {
      const projectTotals = calculateProjectExpenseTotals(projects[projectIndex]);
      const cells = subtotalsRow.querySelectorAll("td");
      if (cells.length >= 4) {
        cells[2].textContent = projectTotals.mileage.toString();
        cells[3].textContent = "$" + projectTotals.expense.toFixed(2);
      }
    }
  }

  // Update grand totals
  updateExpenseTotals(projects);
}

// ===================== EXPENSE SHEET ACTIONS =====================

function openAddExpenseProjectDialog() {
  const dlg = document.getElementById("expenseProjectDlg");
  const list = document.getElementById("expenseProjectList");
  const search = document.getElementById("expenseProjectSearch");

  if (!dlg || !list) return;

  // Get projects from current week's timesheet entries
  const weekKey = formatWeekKey(currentTimesheetWeek);
  const timesheetEntries = getWeekEntries(weekKey);
  const existingExpenseProjects = getWeekExpenses(weekKey);

  // Get unique projects from timesheet rows (by projectId or projectName)
  const timesheetProjects = [];
  const seenKeys = new Map();
  timesheetEntries.forEach((entry) => {
    const projectId = (entry.projectId || "").trim();
    const projectName = (entry.projectName || "").trim();
    if (!projectId && !projectName) return;
    const key = projectId ? `id:${projectId.toLowerCase()}` : `name:${projectName.toLowerCase()}`;
    const existingIndex = seenKeys.get(key);
    if (existingIndex === undefined) {
      seenKeys.set(key, timesheetProjects.length);
      timesheetProjects.push({
        id: projectId,
        name: projectName
      });
      return;
    }
    if (!timesheetProjects[existingIndex].name && projectName) {
      timesheetProjects[existingIndex].name = projectName;
    }
  });

  const renderProjects = (query = "") => {
    list.innerHTML = "";
    const q = query.toLowerCase();

    // Filter timesheet projects by search query
    const filtered = timesheetProjects.filter((p) => {
      const name = (p.name || "").toLowerCase();
      const id = (p.id || "").toLowerCase();
      return !q || name.includes(q) || id.includes(q);
    });

    // Check which projects already have expense entries
    const expenseProjectKeys = new Set(
      existingExpenseProjects.map((p) => {
        const projectId = (p.projectId || "").trim();
        const projectName = (p.projectName || "").trim();
        return projectId ? `id:${projectId.toLowerCase()}` : `name:${projectName.toLowerCase()}`;
      })
    );

    filtered.forEach((project) => {
      const projectKey = project.id
        ? `id:${project.id.toLowerCase()}`
        : `name:${(project.name || "").toLowerCase()}`;
      const alreadyHasExpenses = expenseProjectKeys.has(projectKey);
      const option = el("div", { className: "ts-project-option" + (alreadyHasExpenses ? " disabled" : "") });
      option.innerHTML = `
        <div><strong>${project.id || "--"}</strong> - ${project.name || "Unnamed"}${alreadyHasExpenses ? " (already added)" : ""}</div>
      `;
      if (!alreadyHasExpenses) {
        option.onclick = () => {
          addExpenseProject(project.id, project.name);
          closeDlg("expenseProjectDlg");
        };
      }
      list.appendChild(option);
    });

    if (filtered.length === 0) {
      if (timesheetProjects.length === 0) {
        list.innerHTML = '<p class="muted" style="padding: 1rem; text-align: center;">No projects on timesheet. Add projects to your timesheet first.</p>';
      } else {
        list.innerHTML = '<p class="muted" style="padding: 1rem; text-align: center;">No matching projects found</p>';
      }
    }
  };

  renderProjects();
  if (search) {
    search.value = "";
    search.oninput = () => renderProjects(search.value);
  }

  showDialog(dlg);
}

function addExpenseProject(projectId, projectName) {
  const weekKey = formatWeekKey(currentTimesheetWeek);
  const projects = getWeekExpenses(weekKey);

  // Check if project already exists
  const normalizedId = (projectId || "").trim();
  const normalizedName = (projectName || "").trim();
  if (!normalizedId && !normalizedName) {
    toast("Add a project with a name or job number.");
    return;
  }
  const exists = projects.some((p) => {
    const existingId = (p.projectId || "").trim();
    const existingName = (p.projectName || "").trim();
    if (normalizedId) return existingId === normalizedId;
    if (existingId) return false;
    return existingName.toLowerCase() === normalizedName.toLowerCase();
  });
  if (exists) {
    toast("Project already added. Add entries to the existing project.");
    return;
  }

  const newProject = {
    id: createExpenseProjectId(),
    projectId: normalizedId,
    projectName: normalizedName,
    jobNumber: normalizedId,
    entries: [
      {
        id: createExpenseEntryId(),
        date: "",
        description: "",
        mileage: 0,
        expense: 0
      }
    ],
    images: []
  };

  projects.push(newProject);
  setWeekExpenses(weekKey, projects);
  saveTimesheets();
  renderExpenses();
  toast("Project expense section added.");
}

function removeExpenseProject(projectIndex) {
  if (!confirm("Remove this project and all its expense entries?")) return;

  const weekKey = formatWeekKey(currentTimesheetWeek);
  const projects = getWeekExpenses(weekKey);
  projects.splice(projectIndex, 1);
  setWeekExpenses(weekKey, projects);
  saveTimesheets();
  renderExpenses();
  toast("Project expense section removed.");
}

function openAddExpenseEntryDialog(projectIndex) {
  const dlg = document.getElementById("addExpenseEntryDlg");
  if (!dlg) return;

  document.getElementById("exp_projectIndex").value = projectIndex;
  document.getElementById("exp_date").value = "";
  document.getElementById("exp_description").value = "";
  document.getElementById("exp_mileage").value = "";
  document.getElementById("exp_amount").value = "";

  showDialog(dlg);
}

function saveExpenseEntry() {
  const projectIndex = parseInt(document.getElementById("exp_projectIndex").value);
  const date = document.getElementById("exp_date").value;
  const description = document.getElementById("exp_description").value.trim();
  const mileage = parseFloat(document.getElementById("exp_mileage").value) || 0;
  const expense = parseFloat(document.getElementById("exp_amount").value) || 0;

  const weekKey = formatWeekKey(currentTimesheetWeek);
  const projects = getWeekExpenses(weekKey);

  if (projectIndex < 0 || projectIndex >= projects.length) {
    toast("Invalid project.");
    return;
  }

  const entry = {
    id: createExpenseEntryId(),
    date,
    description,
    mileage,
    expense
  };

  projects[projectIndex].entries.push(entry);
  setWeekExpenses(weekKey, projects);
  saveTimesheets();
  renderExpenses();
  closeDlg("addExpenseEntryDlg");
  toast("Expense entry added.");
}

function removeExpenseEntry(projectIndex, entryIndex) {
  const weekKey = formatWeekKey(currentTimesheetWeek);
  const projects = getWeekExpenses(weekKey);

  if (projectIndex < 0 || projectIndex >= projects.length) return;

  const project = projects[projectIndex];
  if (!project.entries || entryIndex < 0 || entryIndex >= project.entries.length) return;

  project.entries.splice(entryIndex, 1);
  setWeekExpenses(weekKey, projects);
  saveTimesheets();
  renderExpenses();
}

async function addExpenseImages(projectIndex) {
  try {
    const result = await window.pywebview.api.select_expense_images();
    if (result.status === "cancelled" || !result.paths?.length) return;

    const weekKey = formatWeekKey(currentTimesheetWeek);
    const projects = getWeekExpenses(weekKey);

    if (projectIndex < 0 || projectIndex >= projects.length) return;

    const project = projects[projectIndex];
    if (!project.images) project.images = [];

    result.paths.forEach((path) => {
      const filename = path.split(/[\\/]/).pop();
      project.images.push({
        id: createExpenseImageId(),
        path,
        filename
      });
    });

    setWeekExpenses(weekKey, projects);
    saveTimesheets();
    renderExpenses();
    toast(`Added ${result.paths.length} image(s).`);
  } catch (e) {
    console.error("Error selecting images:", e);
    toast("Error selecting images.");
  }
}

function removeExpenseImage(projectIndex, imageIndex) {
  const weekKey = formatWeekKey(currentTimesheetWeek);
  const projects = getWeekExpenses(weekKey);

  if (projectIndex < 0 || projectIndex >= projects.length) return;

  const project = projects[projectIndex];
  if (!project.images || imageIndex < 0 || imageIndex >= project.images.length) return;

  project.images.splice(imageIndex, 1);
  setWeekExpenses(weekKey, projects);
  saveTimesheets();
  renderExpenses();
}

async function exportExpenseSheetToExcel() {
  const weekKey = formatWeekKey(currentTimesheetWeek);
  const projects = getWeekExpenses(weekKey);

  if (!projects.length) {
    toast("No expense entries to export.");
    return;
  }

  try {
    const result = await window.pywebview.api.export_expense_sheet_excel({
      weekKey,
      weekDisplay: formatWeekDisplay(currentTimesheetWeek),
      userName: userSettings.userName || "Employee",
      projects,
      mileageRate: MILEAGE_RATE
    });

    if (result.status === "success") {
      toast("Expense sheet exported successfully.");
    } else {
      throw new Error(result.message);
    }
  } catch (e) {
    console.error("Export failed:", e);
    toast("Failed to export expense sheet: " + e.message);
  }
}

function createIcon(path, size = 16) {
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("viewBox", "0 0 24 24");
  svg.setAttribute("width", String(size));
  svg.setAttribute("height", String(size));
  svg.setAttribute("aria-hidden", "true");
  svg.setAttribute("focusable", "false");
  const p = document.createElementNS("http://www.w3.org/2000/svg", "path");
  p.setAttribute("d", path);
  p.setAttribute("fill", "currentColor");
  svg.appendChild(p);
  return svg;
}

function createIconButton({
  className,
  title,
  ariaLabel,
  path,
  pressed,
  onClick,
}) {
  const props = {
    className,
    type: "button",
    title,
    "aria-label": ariaLabel || title || "",
  };
  if (pressed != null) props["aria-pressed"] = String(pressed);
  const btn = el("button", props);
  btn.appendChild(createIcon(path));
  btn.onclick = (e) => {
    e.stopPropagation();
    if (onClick) onClick(e);
  };
  return btn;
}

function val(id) {
  const el = document.getElementById(id);
  return el ? el.value.trim() : "";
}

function setStat(id, value) {
  const el = document.getElementById(id);
  if (el) el.textContent = value;
}

function closeDlg(id) {
  if (id === "editDlg" && _aiMatchSnapshot) {
    db[_aiMatchSnapshot.index] = normalizeProject(_aiMatchSnapshot.data);
    _aiMatchSnapshot = null;
  }
  document.getElementById(id).close();
}

const debounce = (fn, delay) => {
  let timeoutId;
  return (...args) => {
    clearTimeout(timeoutId);
    timeoutId = setTimeout(() => fn(...args), delay);
  };
};

function parseDueStr(s) {
  if (!s) return null;
  s = s.trim();
  let m = s.match(/^\d{4}-\d{2}-\d{2}$/);
  if (m) return new Date(s + "T12:00:00");
  s = s.replace(/[.]/g, "/").replace(/\s+/g, "");
  const parts = s.split("/");
  if (parts.length === 3) {
    let [mm, dd, yy] = parts;
    if (yy.length === 2) yy = "20" + yy;
    const iso = `${yy}-${mm.padStart(2, "0")}-${dd.padStart(2, "0")}T12:00:00`;
    const d = new Date(iso);
    if (!isNaN(d)) return d;
  }
  const d2 = new Date(s);
  return isNaN(d2) ? null : d2;
}

function dueState(dueStr) {
  const d = parseDueStr(dueStr);
  if (!d) return "ok";
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const dayOfWeek = today.getDay();
  const startOfWeek = new Date(today);
  startOfWeek.setDate(today.getDate() - dayOfWeek);
  const endOfWeek = new Date(startOfWeek);
  endOfWeek.setDate(startOfWeek.getDate() + 6);
  endOfWeek.setHours(23, 59, 59, 999);

  if (d < today) return "overdue";
  if (d >= startOfWeek && d <= endOfWeek) return "dueSoon";
  return "ok";
}

function humanDate(s) {
  const d = parseDueStr(s);
  if (!d) return "";
  return d.toLocaleDateString(undefined, {
    year: "2-digit",
    month: "2-digit",
    day: "2-digit",
  });
}

function formatDueDateShort(date) {
  if (!(date instanceof Date) || isNaN(date)) return "";
  return date.toLocaleDateString(undefined, {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  });
}

// Date validation and formatting functions
function autoFormatDateInput(inputElement) {
  let value = inputElement.value.trim();
  if (!value) return;

  // Remove extra whitespace and normalize separators
  value = value.replace(/\s+/g, '').replace(/[.]/g, '/');

  // Check if it's MM/DD/YY format (2-digit year)
  const parts = value.split('/');
  if (parts.length === 3) {
    let [mm, dd, yy] = parts;

    // If year is 2 digits, expand to 4 digits
    if (yy.length === 2) {
      yy = '20' + yy;
      inputElement.value = `${mm}/${dd}/${yy}`;
    }
  }
}

function validateDueDateInput(inputElement) {
  const value = inputElement.value.trim();
  const validationMsg = inputElement.parentElement.querySelector('.date-validation-msg');

  // Clear previous validation
  inputElement.classList.remove('input-error', 'input-warning');
  if (validationMsg) validationMsg.textContent = '';

  // Empty is valid (optional field)
  if (!value) {
    return true;
  }

  // Check for incomplete date (missing year)
  const parts = value.replace(/[.\s]/g, '/').split('/');
  if (parts.length < 3) {
    inputElement.classList.add('input-error');
    if (validationMsg) {
      validationMsg.textContent = 'Incomplete date. Use MM/DD/YYYY';
      validationMsg.className = 'date-validation-msg error';
    }
    return false;
  }

  // Check if year is present and valid (at least 2 digits)
  const yearPart = parts[2];
  if (!yearPart || yearPart.length < 2) {
    inputElement.classList.add('input-error');
    if (validationMsg) {
      validationMsg.textContent = 'Year required. Use MM/DD/YYYY';
      validationMsg.className = 'date-validation-msg error';
    }
    return false;
  }

  // Try to parse
  const parsed = parseDueStr(value);

  // Invalid format
  if (!parsed) {
    inputElement.classList.add('input-error');
    if (validationMsg) {
      validationMsg.textContent = 'Invalid date format. Use MM/DD/YYYY';
      validationMsg.className = 'date-validation-msg error';
    }
    return false;
  }

  // Valid but in the past
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  if (parsed < today) {
    inputElement.classList.add('input-warning');
    if (validationMsg) {
      validationMsg.textContent = 'Warning: This date is in the past';
      validationMsg.className = 'date-validation-msg warning';
    }
    return true; // Warning, not blocking
  }

  // Valid and future
  return true;
}

function validateAllDueDates() {
  const inputs = document.querySelectorAll('.d-due');
  let allValid = true;

  inputs.forEach(input => {
    if (!validateDueDateInput(input)) {
      allValid = false;
    }
  });

  return allValid;
}

// Calendar utility functions
function getDaysInMonth(year, month) {
  return new Date(year, month + 1, 0).getDate();
}

function getFirstDayOfMonth(year, month) {
  return new Date(year, month, 1).getDay();
}

function isSameDay(date1, date2) {
  return date1.getFullYear() === date2.getFullYear() &&
    date1.getMonth() === date2.getMonth() &&
    date1.getDate() === date2.getDate();
}

function createCalendarGrid(year, month, onDateSelect) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const daysInMonth = getDaysInMonth(year, month);
  const firstDay = getFirstDayOfMonth(year, month);
  const prevMonth = month === 0 ? 11 : month - 1;
  const prevYear = month === 0 ? year - 1 : year;
  const daysInPrevMonth = getDaysInMonth(prevYear, prevMonth);

  const grid = el('div', { className: 'calendar-grid' });

  // Previous month days
  for (let i = firstDay - 1; i >= 0; i--) {
    const day = daysInPrevMonth - i;
    const btn = el('button', {
      className: 'calendar-day other-month',
      textContent: day,
      type: 'button'
    });
    btn.onclick = () => {
      const date = new Date(prevYear, prevMonth, day);
      onDateSelect(date);
    };
    grid.appendChild(btn);
  }

  // Current month days
  for (let day = 1; day <= daysInMonth; day++) {
    const date = new Date(year, month, day);
    const isToday = isSameDay(date, today);
    const classes = ['calendar-day'];
    if (isToday) classes.push('today');

    const btn = el('button', {
      className: classes.join(' '),
      textContent: day,
      type: 'button'
    });
    btn.onclick = () => onDateSelect(date);
    grid.appendChild(btn);
  }

  // Next month days to fill grid
  const totalCells = grid.children.length;
  const remainingCells = totalCells % 7 === 0 ? 0 : 7 - (totalCells % 7);
  const nextMonth = month === 11 ? 0 : month + 1;
  const nextYear = month === 11 ? year + 1 : year;

  for (let day = 1; day <= remainingCells; day++) {
    const btn = el('button', {
      className: 'calendar-day other-month',
      textContent: day,
      type: 'button'
    });
    btn.onclick = () => {
      const date = new Date(nextYear, nextMonth, day);
      onDateSelect(date);
    };
    grid.appendChild(btn);
  }

  return grid;
}

// Inline calendar picker for input fields
function showCalendarForInput(inputElement, onSelectCallback) {
  // Remove any existing calendar
  const existingCalendar = document.getElementById('inlineCalendarPicker');
  if (existingCalendar) existingCalendar.remove();

  // Get current value or default to today
  const currentDate = parseDueStr(inputElement.value) || new Date();

  // Create calendar container
  const calendarContainer = el('div', {
    className: 'inline-calendar-picker',
    id: 'inlineCalendarPicker'
  });

  // Render calendar
  const calendar = renderInlineCalendar(
    currentDate.getFullYear(),
    currentDate.getMonth(),
    (selectedDate) => {
      // On date selection
      inputElement.value = formatDueDateShort(selectedDate);
      // Trigger validation
      validateDueDateInput(inputElement);
      // Callback
      if (onSelectCallback) onSelectCallback(selectedDate);
      // Remove calendar
      calendarContainer.remove();
    },
    () => {
      // On cancel
      calendarContainer.remove();
    }
  );

  calendarContainer.appendChild(calendar);

  // Append inside the closest dialog (top layer) so it's not hidden behind it.
  // Fall back to document.body if not inside a dialog.
  const parentDialog = inputElement.closest('dialog');
  const anchorParent = parentDialog || document.body;
  anchorParent.appendChild(calendarContainer);

  // Position the calendar above the input so it doesn't extend the modal
  const rect = inputElement.getBoundingClientRect();
  const parentRect = anchorParent.getBoundingClientRect();
  calendarContainer.style.position = 'absolute';
  calendarContainer.style.left = `${rect.left - parentRect.left + anchorParent.scrollLeft}px`;
  calendarContainer.style.zIndex = '999999';

  // Measure the calendar's height, then place it above the input
  calendarContainer.style.visibility = 'hidden';
  calendarContainer.style.top = '0px';
  const calHeight = calendarContainer.offsetHeight;
  calendarContainer.style.top = `${rect.top - parentRect.top + anchorParent.scrollTop - calHeight - 5}px`;
  calendarContainer.style.visibility = '';

  // Close on outside click
  setTimeout(() => {
    const closeOnOutsideClick = (e) => {
      if (!calendarContainer.contains(e.target) && e.target !== inputElement) {
        calendarContainer.remove();
        document.removeEventListener('click', closeOnOutsideClick);
      }
    };
    document.addEventListener('click', closeOnOutsideClick);
  }, 100);
}

function showCalendarForDeliverableBadge(anchorElement, deliverable, project) {
  if (!anchorElement || !deliverable) return;

  // Remove any existing calendar
  const existingCalendar = document.getElementById('inlineCalendarPicker');
  if (existingCalendar) existingCalendar.remove();

  // Get current value or default to today
  const currentDate = parseDueStr(deliverable.due) || new Date();

  // Create calendar container
  const calendarContainer = el('div', {
    className: 'inline-calendar-picker',
    id: 'inlineCalendarPicker'
  });

  // Render calendar
  const calendar = renderInlineCalendar(
    currentDate.getFullYear(),
    currentDate.getMonth(),
    async (selectedDate) => {
      deliverable.due = formatDueDateShort(selectedDate);
      if (project) autoSetPrimary(project);
      await save();
      render();
      calendarContainer.remove();
    },
    () => {
      calendarContainer.remove();
    }
  );

  calendarContainer.appendChild(calendar);

  // Append inside the closest dialog (top layer) so it's not hidden behind it.
  // Fall back to document.body if not inside a dialog.
  const parentDialog = anchorElement.closest('dialog');
  const anchorParent = parentDialog || document.body;
  anchorParent.appendChild(calendarContainer);

  // Position the calendar above the anchor
  const rect = anchorElement.getBoundingClientRect();
  const parentRect = anchorParent.getBoundingClientRect();
  calendarContainer.style.position = 'absolute';
  calendarContainer.style.left = `${rect.left - parentRect.left + anchorParent.scrollLeft}px`;
  calendarContainer.style.zIndex = '999999';

  // Measure the calendar's height, then place it above the anchor
  calendarContainer.style.visibility = 'hidden';
  calendarContainer.style.top = '0px';
  const calHeight = calendarContainer.offsetHeight;
  calendarContainer.style.top = `${rect.top - parentRect.top + anchorParent.scrollTop - calHeight - 5}px`;
  calendarContainer.style.visibility = '';

  // Close on outside click
  setTimeout(() => {
    const closeOnOutsideClick = (e) => {
      if (!calendarContainer.contains(e.target) && e.target !== anchorElement) {
        calendarContainer.remove();
        document.removeEventListener('click', closeOnOutsideClick);
      }
    };
    document.addEventListener('click', closeOnOutsideClick);
  }, 100);
}

function renderInlineCalendar(year, month, onDateSelect, onCancel) {
  const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'];

  const container = el('div', { className: 'inline-calendar' });

  // Header with navigation
  const header = el('div', { className: 'calendar-header' });

  const prevBtn = el('button', {
    className: 'calendar-nav-btn',
    textContent: '◀',
    type: 'button'
  });

  const monthYearLabel = el('div', {
    className: 'calendar-month-year',
    textContent: `${monthNames[month]} ${year}`
  });

  const nextBtn = el('button', {
    className: 'calendar-nav-btn',
    textContent: '▶',
    type: 'button'
  });

  header.appendChild(prevBtn);
  header.appendChild(monthYearLabel);
  header.appendChild(nextBtn);
  container.appendChild(header);

  // Weekday labels
  const weekdays = el('div', { className: 'calendar-weekdays' });
  ['Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa'].forEach(day => {
    weekdays.appendChild(el('div', { textContent: day }));
  });
  container.appendChild(weekdays);

  // Calendar grid
  const gridContainer = el('div');
  gridContainer.appendChild(createCalendarGrid(year, month, onDateSelect));
  container.appendChild(gridContainer);

  // Footer with cancel button
  const footer = el('div', { className: 'calendar-footer' });
  const cancelBtn = el('button', {
    className: 'btn ghost',
    textContent: 'Cancel',
    type: 'button'
  });
  cancelBtn.onclick = onCancel;
  footer.appendChild(cancelBtn);
  container.appendChild(footer);

  // Navigation: only swap the grid and update the label to keep DOM stable
  function navigateMonth(newYear, newMonth) {
    monthYearLabel.textContent = `${monthNames[newMonth]} ${newYear}`;
    gridContainer.innerHTML = '';
    gridContainer.appendChild(createCalendarGrid(newYear, newMonth, onDateSelect));

    // Rebind arrows for the new month
    prevBtn.onclick = () => {
      const m = newMonth === 0 ? 11 : newMonth - 1;
      const y = newMonth === 0 ? newYear - 1 : newYear;
      navigateMonth(y, m);
    };
    nextBtn.onclick = () => {
      const m = newMonth === 11 ? 0 : newMonth + 1;
      const y = newMonth === 11 ? newYear + 1 : newYear;
      navigateMonth(y, m);
    };
  }

  prevBtn.onclick = () => {
    const m = month === 0 ? 11 : month - 1;
    const y = month === 0 ? year - 1 : year;
    navigateMonth(y, m);
  };

  nextBtn.onclick = () => {
    const m = month === 11 ? 0 : month + 1;
    const y = month === 11 ? year + 1 : year;
    navigateMonth(y, m);
  };

  return container;
}

// Extract account name from file path
// Handles: P:\Account\..., M:\Account\..., or \\acies.lan\cachedrive\projects2?\Account\...
function extractAccountFromPath(path) {
  if (!path) return null;

  // Normalize path separators to forward slashes
  const normalized = path.replace(/\\/g, '/');

  // Try P: or M: drive pattern first
  let match = normalized.match(/^([PM]):\/?([^\/]+)/i);
  if (match) return match[2];

  // Try network cached drive pattern: \\acies.lan\cachedrive\projects2?\Account\...
  match = normalized.match(/^\/\/[^\/]+\/cachedrive\/projects2?\/([^\/]+)/i);
  if (match) return match[1];

  return null;
}

function basename(path) {
  try {
    if (!path) return "";
    const norm = path.replace(/\\/g, "/");
    const idx = norm.lastIndexOf("/");
    return idx >= 0 ? norm.slice(idx + 1) : norm;
  } catch {
    return path;
  }
}

function toFileURL(raw) {
  if (!raw) return "";
  let s = raw.trim();
  if (/^https?:\/\//i.test(s)) return s;
  if (/^\\\\/.test(s))
    return "file:" + s.replace(/^\\\\/, "/////").replace(/\\/g, "/");
  if (/^[A-Za-z]:\\/.test(s)) return "file:///" + s.replace(/\\/g, "/");
  return s;
}

function normalizeLink(input) {
  const raw = (input || "").trim();
  const url = toFileURL(raw);
  const label = basename(raw) || raw || "link";
  return { label, url, raw };
}

function openExternalUrl(url) {
  try {
    if (window.pywebview?.api?.open_url) {
      window.pywebview.api.open_url(url);
      return;
    }
  } catch {
    /* fallthrough */
  }
  window.open(url, "_blank", "noreferrer");
}

function convertPath(raw) {
  if (raw.startsWith("\\\\acies.lan\\cachedrive\\projects2\\"))
    return raw.replace("\\\\acies.lan\\cachedrive\\projects2\\", "M:\\");
  if (raw.startsWith("\\\\acies.lan\\cachedrive\\projects\\"))
    return raw.replace("\\\\acies.lan\\cachedrive\\projects\\", "P:\\");
  return raw;
}

function parseProjectFromPath(rawPath) {
  if (!rawPath) return null;
  let s = String(rawPath).trim();
  if (!s) return null;
  s = s.replace(/^['"]+|['"]+$/g, "");
  const norm = s.replace(/\\/g, "/").replace(/\/+$/g, "");
  const parts = norm.split("/").filter(Boolean);
  for (let i = parts.length - 1; i >= 0; i--) {
    const segment = parts[i].trim();
    const match = segment.match(/^(\d{5,})\s*(?:[-_]\s*)?(.+)$/);
    if (match) {
      const id = match[1];
      const name = match[2].trim();
      if (name) return { id, name };
    }
  }
  return null;
}

function getProjectBasePath(rawPath) {
  if (!rawPath) return "";
  let s = String(rawPath).trim();
  if (!s) return "";
  s = s.replace(/^['"]+|['"]+$/g, "");
  const norm = s.replace(/\\/g, "/").replace(/\/+$/g, "");
  const parts = norm.split("/").filter(Boolean);
  if (!parts.length) return "";
  let baseIndex = -1;
  for (let i = parts.length - 1; i >= 0; i--) {
    const segment = parts[i].trim();
    if (/^(\d{5,})\s*(?:[-_]\s*)?(.+)$/.test(segment)) {
      baseIndex = i;
      break;
    }
  }
  let baseParts = parts;
  if (baseIndex >= 0) {
    baseParts = parts.slice(0, baseIndex + 1);
  } else {
    const last = parts[parts.length - 1];
    if (/\.[A-Za-z0-9]{1,5}$/.test(last)) baseParts = parts.slice(0, -1);
  }
  return baseParts.join("/");
}

function getProjectBaseKey(rawPath) {
  const base = getProjectBasePath(rawPath);
  return base ? base.toLowerCase() : "";
}

function applyProjectFromPath(path, { force = false } = {}) {
  const parsed = parseProjectFromPath(path);
  if (!parsed) return false;
  const idInput = document.getElementById("f_id");
  const nameInput = document.getElementById("f_name");
  if (!idInput || !nameInput) return false;

  const currentId = idInput.value.trim();
  const currentName = nameInput.value.trim();
  const shouldSetId = force || !currentId;
  const shouldSetName = force || !currentName;

  if (shouldSetId) idInput.value = parsed.id;
  if (shouldSetName) nameInput.value = parsed.name;
  return shouldSetId || shouldSetName;
}

function toast(msg, duration = 2500) {
  const existing = document.querySelector(".toast-notification");
  if (existing) existing.remove();

  const t = el("div", {
    textContent: msg,
    style: `
            position: fixed;
            bottom: 24px;
            left: 50%;
            transform: translateX(-50%);
            background: var(--surface);
            backdrop-filter: blur(12px);
            border: 1px solid var(--accent);
            color: var(--text);
            padding: 0.75rem 1.25rem;
            border-radius: 12px;
            z-index: 9999;
            box-shadow: 0 8px 32px rgba(0,0,0,0.4);
            font-weight: 500;
            display: flex;
            align-items: center;
            gap: 8px;
            animation: slideUp 0.3s ease-out forwards;
        `,
  });

  document.body.append(t);
  setTimeout(() => {
    t.style.opacity = "0";
    t.style.transform = "translateX(-50%) translateY(10px)";
    t.style.transition = "all 0.3s ease";
    setTimeout(() => t.remove(), 300);
  }, duration);
}

function reportClientError(message, error) {
  const detail = error?.message || error?.toString?.() || "";
  const payload = detail ? `${message}: ${detail}` : message;
  console.error(payload, error || "");
  try {
    toast(payload, 6000);
  } catch {
    try {
      alert(payload);
    } catch {
      /* noop */
    }
  }
}

window.addEventListener("error", (event) => {
  reportClientError("JS error", event?.error || event?.message);
});

window.addEventListener("unhandledrejection", (event) => {
  reportClientError("Unhandled promise rejection", event?.reason);
});

function showDialog(dialogEl) {
  if (!dialogEl) return false;
  try {
    dialogEl.showModal();
    return true;
  } catch (err) {
    try {
      dialogEl.setAttribute("open", "");
      return true;
    } catch (fallbackErr) {
      reportClientError("Failed to open dialog", fallbackErr);
      return false;
    }
  }
}

const updateStickyOffsets = () => {
  const header = document.querySelector(".app-header");
  const toolbar = document.querySelector("#projects-panel .panel-toolbar");
  if (header)
    document.documentElement.style.setProperty(
      "--header-height",
      `${header.offsetHeight}px`
    );
  if (toolbar)
    document.documentElement.style.setProperty(
      "--toolbar-height",
      `${toolbar.offsetHeight}px`
    );
};

const debouncedStickyOffsets = debounce(updateStickyOffsets, 150);
window.addEventListener("resize", debouncedStickyOffsets);

// ===================== SERVER I/O =====================

// State variables
let db = [];
let notesDb = {};
let noteTabs = [];
let editIndex = -1;
let _aiMatchSnapshot = null;
let currentSort = { key: "due", dir: "desc" };
let statusFilter = "all";
let dueFilter = "all";

const DEFAULT_CLEAN_DWG_OPTIONS = {
  stripXrefs: true,
  setByLayer: true,
  purge: true,
  audit: true,
  hatchColor: true,
};
const DEFAULT_PUBLISH_DWG_OPTIONS = {
  autoDetectPaperSize: true,
  shrinkPercent: 100,
};
const DEFAULT_FREEZE_LAYER_OPTIONS = {
  scanAllLayers: true,
};
const DEFAULT_THAW_LAYER_OPTIONS = {
  scanAllLayers: true,
};

let userSettings = {
  userName: "",
  discipline: ["Electrical"],
  apiKey: "",
  autocadPath: "",
  showSetupHelp: true,
  theme: "dark",
  lightingTemplates: [],
  autoPrimary: false,
  defaultPmInitials: "",
  cleanDwgOptions: { ...DEFAULT_CLEAN_DWG_OPTIONS },
  publishDwgOptions: { ...DEFAULT_PUBLISH_DWG_OPTIONS },
  freezeLayerOptions: { ...DEFAULT_FREEZE_LAYER_OPTIONS },
  thawLayerOptions: { ...DEFAULT_THAW_LAYER_OPTIONS },
};
let hideNonPrimary = true;
let activeNoteTab = null;
let latestAppUpdate = null;
let currentStatsTimespan = "1Y";
let currentStatsAggregation = "month";
let lightingScheduleProjectIndex = null;
let lightingScheduleProjectQuery = "";
let lightingTemplateQuery = "";

// Timesheet State
let timesheetDb = { weeks: {}, lastModified: null };
let currentTimesheetWeek = getWeekStartDate(new Date());

// Templates State
let templatesDb = { templates: [], defaultTemplatesInstalled: false, lastModified: null };
const DISCIPLINE_OPTIONS = ['General', 'Electrical', 'Mechanical', 'Plumbing'];
const FILE_TYPE_ICONS = { doc: '📄', docx: '📄', dwg: '📐', xlsx: '📊', xls: '📊' };
const FILE_TYPE_LABELS = { doc: 'Word Document', docx: 'Word Document', dwg: 'AutoCAD Drawing', xlsx: 'Excel Spreadsheet', xls: 'Excel Spreadsheet' };
const TEMPLATE_KEY_BY_NAME = {
  "narrative of changes": "narrative",
  "plan check comments": "planCheck",
  "plan check response letter": "planCheck",
};

let copyDialogTemplate = null;
let copyDialogTemplateKey = null;
let copyDialogDeliverables = [];
let copyDialogAutoError = "";

const allowedAggregations = {
  "1W": ["day", "week"],
  "1M": ["day", "week", "month"],
  "1Y": ["week", "month", "quarter"],
  "3Y": ["month", "quarter"],
};

// ===================== THEMING =====================
function readLocalTheme() {
  try {
    return localStorage.getItem(THEME_STORAGE_KEY);
  } catch (e) {
    return null;
  }
}

function applyTheme(theme) {
  const resolved = theme === "light" ? "light" : "dark";
  document.documentElement.setAttribute("data-theme", resolved);
  const toggle = document.getElementById("themeToggleBtn");
  if (toggle) {
    const isLight = resolved === "light";
    const label = isLight ? "Switch to dark mode" : "Switch to light mode";
    toggle.textContent = isLight ? "🌙" : "☀️";
    toggle.setAttribute("aria-label", label);
    toggle.title = label;
  }
  try {
    localStorage.setItem(THEME_STORAGE_KEY, resolved);
  } catch (e) {
    // Ignore storage errors (private mode, etc.)
  }
  userSettings.theme = resolved;
  return resolved;
}

async function persistThemePreference(theme) {
  const resolved = applyTheme(theme);
  if (window.pywebview && window.pywebview.api?.save_user_settings) {
    try {
      await window.pywebview.api.save_user_settings(userSettings);
    } catch (e) {
      console.warn("Failed to persist theme preference:", e);
    }
  }
  return resolved;
}

function initThemeFromPreferences() {
  const storedSetting = userSettings.theme;
  const localPref = readLocalTheme();
  const prefersDark = window.matchMedia("(prefers-color-scheme: dark)").matches;
  const nextTheme =
    storedSetting || localPref || (prefersDark ? "dark" : "light");
  applyTheme(nextTheme);
}

async function load() {
  try {
    const arr = await window.pywebview.api.get_tasks();
    const { data, didMigrate } = migrateProjects(arr);
    migrateStatuses(data);
    if (didMigrate) {
      db = data;
      await save();
    }
    return data;
  } catch (e) {
    console.warn("Backend load failed:", e);
    return [];
  }
}

async function save() {
  try {
    const response = await window.pywebview.api.save_tasks(db);
    if (response.status !== "success") throw new Error(response.message);
  } catch (e) {
    console.warn("Backend save failed:", e);
    toast("⚠️ Failed to save data.");
  }
}

async function refreshAppUpdateStatus({ manual = false } = {}) {
  const versionLabel = document.getElementById("appVersionLabel");
  const versionChip = document.getElementById("versionChip");
  const updateBtn = document.getElementById("appUpdateBtn");

  try {
    const res = await window.pywebview.api.get_app_update_status();
    if (versionLabel && res.current_version) {
      versionLabel.textContent = `v${res.current_version}`;
    }

    if (res.status !== "success")
      throw new Error(res.message || "Update check failed");

    if (res.update_available) {
      latestAppUpdate = res;
      if (updateBtn) {
        updateBtn.style.display = "inline-flex";
        updateBtn.textContent = `Update to v${res.latest_version}`;
        updateBtn.dataset.downloadUrl = res.download_url || "";
      }
      versionChip?.classList.add("has-update");
      if (manual) toast(`Update available (v${res.latest_version}).`);
    } else {
      latestAppUpdate = null;
      if (updateBtn) {
        updateBtn.style.display = "none";
        updateBtn.removeAttribute("data-download-url");
        updateBtn.textContent = "Update";
      }
      versionChip?.classList.remove("has-update");
      if (manual) toast("You are on the latest version.");
    }
  } catch (e) {
    console.warn("Update check failed:", e);
    if (manual) toast("Update check failed.");
  }
}

async function installAppUpdate() {
  const updateBtn = document.getElementById("appUpdateBtn");
  const downloadUrl =
    updateBtn?.dataset.downloadUrl || latestAppUpdate?.download_url;
  if (!downloadUrl) {
    toast("No update available.");
    return;
  }

  try {
    if (updateBtn) {
      updateBtn.disabled = true;
      updateBtn.textContent = "Downloading...";
    }
    toast("Updating... the app will restart when finished.");
    const res = await window.pywebview.api.download_and_install_app_update(
      downloadUrl
    );
    if (res.status !== "success") throw new Error(res.message);
    toast("Installer is running. The app will reopen after the update.");
  } catch (e) {
    console.warn("Update install failed:", e);
    toast(`Update failed: ${e.message || e}`);
  } finally {
    if (updateBtn) {
      updateBtn.disabled = false;
      const label = latestAppUpdate?.latest_version
        ? `Update to v${latestAppUpdate.latest_version}`
        : "Update";
      updateBtn.textContent = label;
      if (!latestAppUpdate) updateBtn.style.display = "none";
    }
  }
}

async function loadNotes() {
  try {
    const data = (await window.pywebview.api.get_notes()) || {};
    noteTabs =
      Array.isArray(data.tabs) && data.tabs.length > 0
        ? data.tabs
        : ["General"];
    notesDb = {};
    noteTabs.forEach((tab) => {
      const keyedContent =
        data.keyed && data.keyed[tab] ? data.keyed[tab].trim() : "";
      const generalContent =
        data.general && data.general[tab] ? data.general[tab].trim() : "";
      if (keyedContent && generalContent) {
        notesDb[
          tab
        ] = `${keyedContent}\n\n--- General Notes ---\n\n${generalContent}`;
      } else {
        notesDb[tab] = keyedContent || generalContent;
      }
    });
    activeNoteTab = noteTabs[0];
    return notesDb;
  } catch (e) {
    noteTabs = ["General"];
    notesDb = {};
    activeNoteTab = noteTabs[0];
    return {};
  }
}

async function saveNotes() {
  try {
    const dataToSave = { tabs: noteTabs, keyed: {}, general: notesDb };
    await window.pywebview.api.save_notes(dataToSave);
  } catch (e) {
    console.warn("Backend notes save failed:", e);
  }
}

async function loadUserSettings() {
  try {
    const storedSettings = await window.pywebview.api.get_user_settings();
    if (storedSettings) userSettings = { ...userSettings, ...storedSettings };
    userSettings.discipline = normalizeDisciplineList(userSettings.discipline);
    userSettings.cleanDwgOptions = {
      ...DEFAULT_CLEAN_DWG_OPTIONS,
      ...(userSettings.cleanDwgOptions || {}),
    };
    userSettings.publishDwgOptions = {
      ...DEFAULT_PUBLISH_DWG_OPTIONS,
      ...(userSettings.publishDwgOptions || {}),
    };
    userSettings.freezeLayerOptions = {
      ...DEFAULT_FREEZE_LAYER_OPTIONS,
      ...(userSettings.freezeLayerOptions || {}),
    };
    userSettings.thawLayerOptions = {
      ...DEFAULT_THAW_LAYER_OPTIONS,
      ...(userSettings.thawLayerOptions || {}),
    };
  } catch (e) {
    console.error("Failed to load settings:", e);
  }
}

function setCheckboxValue(id, value) {
  const checkbox = document.getElementById(id);
  if (checkbox) checkbox.checked = !!value;
}

function setRangeValue(id, value) {
  const slider = document.getElementById(id);
  if (!slider) return;
  const numeric = Number(value);
  slider.value = Number.isFinite(numeric) ? String(numeric) : slider.value;
}

function setTextValue(id, value) {
  const el = document.getElementById(id);
  if (el) el.textContent = value;
}

function syncCleanOptionsInputs() {
  const cleanOptions = userSettings.cleanDwgOptions || {};
  setCheckboxValue("settings_clean_stripXrefs", cleanOptions.stripXrefs);
  setCheckboxValue("settings_clean_setByLayer", cleanOptions.setByLayer);
  setCheckboxValue("settings_clean_purge", cleanOptions.purge);
  setCheckboxValue("settings_clean_audit", cleanOptions.audit);
  setCheckboxValue("settings_clean_hatchColor", cleanOptions.hatchColor);
  setCheckboxValue("clean_modal_stripXrefs", cleanOptions.stripXrefs);
  setCheckboxValue("clean_modal_setByLayer", cleanOptions.setByLayer);
  setCheckboxValue("clean_modal_purge", cleanOptions.purge);
  setCheckboxValue("clean_modal_audit", cleanOptions.audit);
  setCheckboxValue("clean_modal_hatchColor", cleanOptions.hatchColor);
}

function syncPublishOptionsInputs() {
  const publishOptions = userSettings.publishDwgOptions || {};
  setCheckboxValue(
    "settings_publish_autoDetectPaperSize",
    publishOptions.autoDetectPaperSize
  );
  setCheckboxValue(
    "publish_modal_autoDetectPaperSize",
    publishOptions.autoDetectPaperSize
  );
  const percent = Number(publishOptions.shrinkPercent ?? 100);
  const normalized = Number.isFinite(percent) ? percent : 100;
  setRangeValue("settings_publish_shrinkPercent", normalized);
  setRangeValue("publish_modal_shrinkPercent", normalized);
  setTextValue("settings_publish_shrinkValue", `${normalized}%`);
  setTextValue("publish_modal_shrinkValue", `${normalized}%`);
}

function syncFreezeOptionsInputs() {
  const freezeOptions = userSettings.freezeLayerOptions || {};
  setCheckboxValue(
    "settings_freeze_scanAllLayers",
    freezeOptions.scanAllLayers
  );
  setCheckboxValue("freeze_modal_scanAllLayers", freezeOptions.scanAllLayers);
}

function syncThawOptionsInputs() {
  const thawOptions = userSettings.thawLayerOptions || {};
  setCheckboxValue("settings_thaw_scanAllLayers", thawOptions.scanAllLayers);
  setCheckboxValue("thaw_modal_scanAllLayers", thawOptions.scanAllLayers);
}

async function populateSettingsModal() {
  document.getElementById("settings_userName").value =
    userSettings.userName || "";
  document.getElementById("settings_apiKey").value = userSettings.apiKey || "";
  document.getElementById("settings_defaultPmInitials").value =
    userSettings.defaultPmInitials || "";
  document.getElementById("settings_autocadPath").value =
    userSettings.autocadPath || "";
  const disciplines = normalizeDisciplineList(userSettings.discipline);
  document
    .querySelectorAll('input[name="settings_discipline_checkbox"]')
    .forEach((checkbox) => {
      checkbox.checked = disciplines.includes(checkbox.value);
    });

  const autoPrimaryCheck = document.getElementById("settings_autoPrimary");
  if (autoPrimaryCheck) autoPrimaryCheck.checked = !!userSettings.autoPrimary;

  syncCleanOptionsInputs();
  syncPublishOptionsInputs();
  syncFreezeOptionsInputs();
  syncThawOptionsInputs();

  await refreshTimesheetsInfo();

  // Populate AutoCAD versions
  const container = document.getElementById("autocad_versions_container");
  container.innerHTML = '<div class="spinner">Detecting versions...</div>';
  try {
    const response =
      await window.pywebview.api.get_installed_autocad_versions();
    if (response.status === "success") {
      container.innerHTML = "";
      response.versions.forEach((version) => {
        const radio = el("input", {
          type: "radio",
          name: "autocad_version_radio",
          value: version.path,
          checked: userSettings.autocadPath === version.path,
        });
        radio.onchange = () => {
          userSettings.autocadPath = radio.value;
          debouncedSaveUserSettings();
        };
        const label = el("label", { className: "radio-label" }, [
          radio,
          ` AutoCAD ${version.year}`,
        ]);
        container.appendChild(label);
      });
      if (response.versions.length === 0) {
        container.innerHTML =
          '<p class="tiny muted">No AutoCAD versions detected in default location.</p>';
      }
    } else {
      container.innerHTML =
        '<p class="tiny muted">Error detecting versions.</p>';
    }
  } catch (e) {
    container.innerHTML = '<p class="tiny muted">Error detecting versions.</p>';
  }
}

async function populateAutocadSelectModal() {
  document.getElementById("autocad_select_custom").value =
    userSettings.autocadPath || "";
  const container = document.getElementById("autocad_select_container");
  container.innerHTML = '<div class="spinner">Detecting versions...</div>';
  try {
    const response =
      await window.pywebview.api.get_installed_autocad_versions();
    if (response.status === "success") {
      container.innerHTML = "";
      response.versions.forEach((version) => {
        const radio = el("input", {
          type: "radio",
          name: "autocad_select_radio",
          value: version.path,
          checked: userSettings.autocadPath === version.path,
        });
        const label = el("label", { className: "radio-label" }, [
          radio,
          ` AutoCAD ${version.year}`,
        ]);
        container.appendChild(label);
      });
      if (response.versions.length === 0) {
        container.innerHTML =
          '<p class="tiny muted">No AutoCAD versions detected in default location.</p>';
      }
    } else {
      container.innerHTML =
        '<p class="tiny muted">Error detecting versions.</p>';
    }
  } catch (e) {
    container.innerHTML = '<p class="tiny muted">Error detecting versions.</p>';
  }
}

async function showAutocadSelectModal() {
  await populateAutocadSelectModal();
  document.getElementById("autocadSelectDlg").showModal();
}

async function saveUserSettings() {
  // Update autocadPath from UI
  const selectedRadio = document.querySelector(
    'input[name="autocad_version_radio"]:checked'
  );
  if (selectedRadio) {
    userSettings.autocadPath = selectedRadio.value;
  } else {
    userSettings.autocadPath = document
      .getElementById("settings_autocadPath")
      .value.trim();
  }
  const autoPrimaryCheck = document.getElementById("settings_autoPrimary");
  if (autoPrimaryCheck) userSettings.autoPrimary = autoPrimaryCheck.checked;
  try {
    await window.pywebview.api.save_user_settings(userSettings);
  } catch (e) {
    toast("⚠️ Could not save settings.");
  }
}
const debouncedSaveUserSettings = debounce(saveUserSettings, 500);

function createId(prefix = "dlv") {
  if (window.crypto?.randomUUID) return `${prefix}_${crypto.randomUUID()}`;
  return `${prefix}_${Date.now().toString(36)}_${Math.random()
    .toString(36)
    .slice(2, 8)}`;
}

function normalizeTaskLinks(links) {
  if (!Array.isArray(links)) return [];
  return links
    .map((link) => {
      if (!link) return null;
      if (typeof link === "string") return normalizeLink(link);
      const raw = link.raw || link.url || "";
      if (!raw) return null;
      const normalized = normalizeLink(raw);
      if (link.label) normalized.label = link.label;
      return normalized;
    })
    .filter(Boolean);
}

function normalizeTask(task) {
  if (typeof task === "string")
    return { text: task, done: false, links: [] };
  return {
    text: task?.text || "",
    done: !!task?.done,
    links: normalizeTaskLinks(task?.links),
  };
}

const DELIVERABLE_NAME_RULES = [
  { regex: /\bDD\s*50\b/i, label: "DD50" },
  { regex: /\bDD\s*90\b/i, label: "DD90" },
  { regex: /\bCD\s*50\b/i, label: "CD50" },
  { regex: /\bCD\s*90\b/i, label: "CD90" },
  { regex: /\bCDF\b/i, label: "CDF" },
  { regex: /\bIFP\b/i, label: "IFP" },
  { regex: /\bIFC\b/i, label: "IFC" },
  { regex: /\bASR\b/i, label: "ASR" },
];

function extractDeliverableName(text) {
  if (!text) return "";
  const raw = String(text);
  const rfiMatch = raw.match(/\bRFI\s*#?\s*(\d+)\b/i);
  if (rfiMatch) return `RFI #${rfiMatch[1]}`;
  if (/\bRFI\b/i.test(raw)) return "RFI";
  const pccMatch = raw.match(/\bPCC\s*#?\s*(\d+)\b/i);
  if (pccMatch) return `PCC ${pccMatch[1]}`;
  if (/\bPCC\b/i.test(raw)) return "PCC";
  const asrMatch = raw.match(/\bASR\s*#?\s*(\d+)\b/i);
  if (asrMatch) return `ASR #${asrMatch[1]}`;
  for (const rule of DELIVERABLE_NAME_RULES) {
    if (rule.regex.test(raw)) return rule.label;
  }
  return "";
}

function guessDeliverableName(legacy) {
  const candidates = [];
  if (legacy?.deliverable) candidates.push(legacy.deliverable);
  if (legacy?.nick) candidates.push(legacy.nick);
  if (legacy?.name) candidates.push(legacy.name);
  if (legacy?.path) candidates.push(basename(legacy.path));
  if (Array.isArray(legacy?.tasks)) {
    legacy.tasks.forEach((t) => candidates.push(t?.text || t));
  }
  if (legacy?.notes) candidates.push(legacy.notes);
  for (const text of candidates) {
    const name = extractDeliverableName(text);
    if (name) return name;
  }
  return "Deliverable";
}

function normalizeDeliverable(deliverable = {}) {
  const out = {
    id: deliverable.id || createId("dlv"),
    name: String(deliverable.name || "").trim(),
    due: String(deliverable.due || "").trim(),
    notes: deliverable.notes || "",
    tasks: Array.isArray(deliverable.tasks)
      ? deliverable.tasks.map(normalizeTask)
      : [],
    statuses: Array.isArray(deliverable.statuses)
      ? [...deliverable.statuses]
      : [],
    statusTags: Array.isArray(deliverable.statusTags)
      ? [...deliverable.statusTags]
      : [],
    status: deliverable.status || "",
  };
  migrateStatusFields(out);
  syncStatusArrays(out);
  if (isFinished(out)) {
    out.tasks.forEach((task) => {
      task.done = true;
    });
  }
  return out;
}

function createDeliverable(seed = {}) {
  return normalizeDeliverable({
    id: seed.id || createId("dlv"),
    name: seed.name || "",
    due: seed.due || "",
    notes: seed.notes || "",
    tasks: seed.tasks || [],
    statuses: seed.statuses || [],
    statusTags: seed.statusTags || [],
    status: seed.status || "",
  });
}

// ===================== DATA MIGRATION =====================
function canonStatus(s) {
  if (!s) return null;
  const t = String(s).trim().toLowerCase();
  if (["waiting", "wait", "blocked"].includes(t)) return "Waiting";
  if (
    [
      "working",
      "work",
      "in progress",
      "in-progress",
      "doing",
      "active",
    ].includes(t)
  )
    return "Working";
  if (
    ["pending review", "pending-review", "review", "pr", "pending"].includes(t)
  )
    return "Pending Review";
  if (["complete", "completed", "done"].includes(t)) return "Complete";
  if (["delivered", "sent", "shipped"].includes(t)) return "Delivered";
  return null;
}
function migrateStatusFields(item) {
  if (!item) return;
  if (!Array.isArray(item.statuses)) item.statuses = [];
  if (item.status) {
    String(item.status)
      .split(/[,/|;]+/)
      .map((s) => s.trim())
      .filter(Boolean)
      .forEach((piece) => {
        const c = canonStatus(piece);
        if (c && !item.statuses.includes(c)) item.statuses.push(c);
      });
  }
  item.statuses = item.statuses
    .map((s) => canonStatus(s) || s)
    .filter((s) => STATUS_CANON.includes(s));
}
function hasStatus(p, s) {
  return Array.isArray(p.statuses) && p.statuses.includes(s);
}
function isFinished(p) {
  return hasStatus(p, "Complete") || hasStatus(p, "Delivered");
}
function applyPrimaryStatus(p) {
  const primary = STATUS_PRIORITY.find((status) => hasStatus(p, status));
  p.status = primary || "";
}
function toggleStatus(p, label) {
  if (!Array.isArray(p.statuses)) p.statuses = [];
  const key = LABEL_TO_KEY[label];
  if (key) setTag(p, key, !p.statuses.includes(label));
  if (isFinished(p) && Array.isArray(p.tasks)) {
    p.tasks.forEach((task) => {
      task.done = true;
    });
  }
}
function syncStatusArrays(p) {
  if (!Array.isArray(p.statuses)) p.statuses = [];
  const fromTags = Array.isArray(p.statusTags) ? p.statusTags : [];
  for (const k of fromTags) {
    const L = KEY_TO_LABEL[k];
    if (L && !p.statuses.includes(L)) p.statuses.push(L);
  }
  p.statuses = [...new Set(p.statuses.filter((s) => STATUS_CANON.includes(s)))];
  p.statusTags = p.statuses.map((s) => LABEL_TO_KEY[s]).filter(Boolean);
  applyPrimaryStatus(p);
}
function migrateStatuses(arr) {
  for (const p of arr) {
    if (Array.isArray(p.deliverables)) {
      p.deliverables.forEach((d) => {
        migrateStatusFields(d);
        syncStatusArrays(d);
      });
    } else {
      migrateStatusFields(p);
      syncStatusArrays(p);
    }
  }
}
function setTag(p, key, on) {
  const tags = getStatusTags(p);
  const idx = tags.indexOf(key);
  if (on && idx === -1) tags.push(key);
  if (!on && idx !== -1) tags.splice(idx, 1);
  p.statusTags = tags;
  const label = KEY_TO_LABEL[key];
  if (!Array.isArray(p.statuses)) p.statuses = [];
  const j = p.statuses.indexOf(label);
  if (label) {
    if (on && j === -1) p.statuses.push(label);
    if (!on && j !== -1) p.statuses.splice(j, 1);
  }
  applyPrimaryStatus(p);
}
function getStatusTags(p) {
  let tags = Array.isArray(p.statusTags) ? [...p.statusTags] : [];
  const s = (p.status || "").toLowerCase();
  if (s) {
    if (s.includes("complete") && !tags.includes("complete"))
      tags.push("complete");
    if (s.includes("waiting") && !tags.includes("waiting"))
      tags.push("waiting");
    if (
      (s.includes("working") || s.includes("in progress")) &&
      !tags.includes("working")
    )
      tags.push("working");
    if (
      (s.includes("pending review") || s === "pending") &&
      !tags.includes("pendingReview")
    )
      tags.push("pendingReview");
    if (s.includes("deliver") && !tags.includes("delivered"))
      tags.push("delivered");
  }
  return [...new Set(tags)];
}

function normalizeRef(link) {
  if (!link) return null;
  if (typeof link === "string") return normalizeLink(link);
  const raw = link.raw || link.url || "";
  if (!raw) return null;
  const normalized = normalizeLink(raw);
  if (link.label) normalized.label = link.label;
  return normalized;
}

function isLegacyProject(project) {
  if (!project) return false;
  if (Array.isArray(project.deliverables)) return false;
  return (
    project.due ||
    project.tasks ||
    project.statuses ||
    project.status ||
    project.notes
  );
}

function getProjectMergeKey(project, index) {
  const id = String(project?.id || "").trim().toLowerCase();
  if (id) return `id:${id}`;
  const path = String(project?.path || "").trim().toLowerCase();
  if (path) return `path:${path}`;
  const name = String(project?.name || "").trim().toLowerCase();
  if (name) return `name:${name}`;
  return `__project_${index}`;
}

function mergeRefs(baseRefs = [], incomingRefs = []) {
  const out = [];
  const seen = new Set();
  [...baseRefs, ...incomingRefs].forEach((ref) => {
    const normalized = normalizeRef(ref);
    if (!normalized) return;
    const key = (normalized.raw || normalized.url || "").toLowerCase();
    if (!key || seen.has(key)) return;
    seen.add(key);
    out.push(normalized);
  });
  return out;
}

function mergeProjects(base, incoming) {
  if (!base || !incoming) return;
  if (!base.id && incoming.id) base.id = incoming.id;
  if (!base.name && incoming.name) base.name = incoming.name;
  if (!base.nick && incoming.nick) base.nick = incoming.nick;
  if (!base.path && incoming.path) base.path = incoming.path;
  if (!base.notes && incoming.notes) base.notes = incoming.notes;
  base.refs = mergeRefs(base.refs || [], incoming.refs || []);
  if (!base.lightingSchedule && incoming.lightingSchedule)
    base.lightingSchedule = incoming.lightingSchedule;
  if (!Array.isArray(base.deliverables)) base.deliverables = [];
  if (Array.isArray(incoming.deliverables))
    base.deliverables.push(...incoming.deliverables);
  if (!base.overviewDeliverableId && incoming.overviewDeliverableId)
    base.overviewDeliverableId = incoming.overviewDeliverableId;
}

function convertLegacyProject(legacy) {
  const deliverable = createDeliverable({
    name: guessDeliverableName(legacy),
    due: legacy?.due || "",
    notes: legacy?.notes || "",
    tasks: legacy?.tasks || [],
    statuses: legacy?.statuses || [],
    statusTags: legacy?.statusTags || [],
    status: legacy?.status || "",
  });
  return {
    id: String(legacy?.id || "").trim(),
    name: String(legacy?.name || "").trim(),
    nick: String(legacy?.nick || "").trim(),
    path: String(legacy?.path || "").trim(),
    notes: "",
    refs: Array.isArray(legacy?.refs)
      ? legacy.refs.map(normalizeRef).filter(Boolean)
      : [],
    deliverables: [deliverable],
    overviewDeliverableId: deliverable.id,
    lightingSchedule: legacy?.lightingSchedule || null,
  };
}

function normalizeProject(project) {
  if (!project) return null;
  if (!Array.isArray(project.deliverables) && isLegacyProject(project)) {
    return normalizeProject(convertLegacyProject(project));
  }
  const out = {
    ...project,
    id: String(project.id || "").trim(),
    name: String(project.name || "").trim(),
    nick: String(project.nick || "").trim(),
    path: String(project.path || "").trim(),
    notes: project.notes || "",
    refs: Array.isArray(project.refs)
      ? project.refs.map(normalizeRef).filter(Boolean)
      : [],
    deliverables: Array.isArray(project.deliverables)
      ? project.deliverables.map(normalizeDeliverable)
      : [],
  };
  if (!out.deliverables.length) out.deliverables = [createDeliverable()];
  if (
    !out.overviewDeliverableId ||
    !out.deliverables.some((d) => d.id === out.overviewDeliverableId)
  ) {
    out.overviewDeliverableId = out.deliverables[0]?.id || "";
  }
  return out;
}

function getLatestDueDeliverableId(deliverables = []) {
  let latestId = "";
  let latestDue = null;
  deliverables.forEach((deliverable) => {
    const due = parseDueStr(deliverable?.due);
    if (!due) return;
    if (!latestDue || due > latestDue) {
      latestDue = due;
      latestId = deliverable?.id || "";
    }
  });
  return latestId;
}

function autoSetPrimary(project) {
  if (!project || !userSettings.autoPrimary) return;
  const latestId = getLatestDueDeliverableId(project.deliverables);
  if (latestId && project.overviewDeliverableId !== latestId) {
    project.overviewDeliverableId = latestId;
  }
}

function migrateProjects(raw = []) {
  let changed = false;
  const map = new Map();
  const legacyKeys = new Set();
  const nonLegacyKeys = new Set();
  raw.forEach((item, index) => {
    const isLegacy = isLegacyProject(item);
    let project = item;
    if (isLegacy) {
      project = convertLegacyProject(item);
      changed = true;
    } else {
      project = normalizeProject(item);
    }
    const key = getProjectMergeKey(project, index);
    if (isLegacy) {
      legacyKeys.add(key);
    } else {
      nonLegacyKeys.add(key);
    }
    if (map.has(key)) {
      mergeProjects(map.get(key), project);
      changed = true;
    } else {
      map.set(key, project);
    }
  });
  const legacyOnlyKeys = new Set(
    [...legacyKeys].filter((key) => !nonLegacyKeys.has(key))
  );
  const merged = Array.from(map.entries())
    .map(([key, p]) => {
      const normalized = normalizeProject(p);
      if (normalized && legacyOnlyKeys.has(key)) {
        const latestDueId = getLatestDueDeliverableId(
          normalized.deliverables
        );
        if (
          latestDueId &&
          normalized.overviewDeliverableId !== latestDueId
        ) {
          normalized.overviewDeliverableId = latestDueId;
          changed = true;
        }
      }
      return normalized;
    })
    .filter(Boolean);
  return { data: merged, didMigrate: changed };
}

function getProjectDeliverables(project) {
  return Array.isArray(project?.deliverables) ? project.deliverables : [];
}

function getAllDeliverables() {
  return db.flatMap((project) =>
    getProjectDeliverables(project).map((deliverable) => ({
      project,
      deliverable,
    }))
  );
}

function compareDeliverablesByDue(a, b) {
  const da = parseDueStr(a?.due);
  const dbb = parseDueStr(b?.due);
  if (!da && !dbb) return 0;
  if (!da) return 1;
  if (!dbb) return -1;
  return da - dbb;
}

function compareDeliverablesByDueDesc(a, b) {
  const da = parseDueStr(a?.due);
  const dbb = parseDueStr(b?.due);
  if (!da && !dbb) return 0;
  if (!da) return 1;
  if (!dbb) return -1;
  return dbb - da;
}

function sortDeliverablesByPrimaryThenDueDesc(list, primaryId) {
  list.sort((a, b) => {
    const aPrimary = a?.id === primaryId;
    const bPrimary = b?.id === primaryId;
    if (aPrimary && !bPrimary) return -1;
    if (!aPrimary && bPrimary) return 1;
    return compareDeliverablesByDueDesc(a, b);
  });
}

function getEarliestIncompleteDeliverable(project) {
  const deliverables = getProjectDeliverables(project).filter(
    (d) => !isFinished(d)
  );
  if (!deliverables.length) return null;
  const withDue = deliverables.filter((d) => parseDueStr(d?.due));
  if (withDue.length) return withDue.sort(compareDeliverablesByDue)[0];
  return deliverables[0];
}

function getPriorityDeliverable(project) {
  const deliverables = getProjectDeliverables(project);
  if (!deliverables.length) return null;
  const priorityId = project?.overviewDeliverableId;
  if (priorityId) {
    const priority = deliverables.find((d) => d.id === priorityId);
    if (priority) return priority;
  }
  return getEarliestIncompleteDeliverable(project) || deliverables[0];
}

function getPrimaryDeliverable(project) {
  const deliverables = getProjectDeliverables(project);
  if (!deliverables.length) return null;
  const primaryId = project?.overviewDeliverableId;
  if (primaryId) {
    const primary = deliverables.find((d) => d.id === primaryId);
    if (primary) return primary;
  }
  return deliverables[0];
}

function getOverviewDeliverables(project) {
  const deliverables = getProjectDeliverables(project);
  if (!deliverables.length) return [];
  const out = deliverables.slice();
  sortDeliverablesByPrimaryThenDueDesc(out, project?.overviewDeliverableId);
  return out;
}

function getProjectSortKey(project) {
  const primary = getPrimaryDeliverable(project);
  return parseDueStr(primary?.due);
}

function matchesDueFilter(deliverable, filter) {
  if (filter === "all") return true;
  const d = parseDueStr(deliverable?.due);
  if (!d) return false;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const dayOfWeek = today.getDay();
  const startOfWeek = new Date(today);
  startOfWeek.setDate(today.getDate() - dayOfWeek);
  startOfWeek.setHours(0, 0, 0, 0);
  const endOfWeek = new Date(startOfWeek);
  endOfWeek.setDate(startOfWeek.getDate() + 6);
  endOfWeek.setHours(23, 59, 59, 999);

  if (filter === "past") return d < startOfWeek;
  if (filter === "soon") return d >= startOfWeek && d <= endOfWeek;
  if (filter === "future") return d > endOfWeek;
  return true;
}

// ===================== RENDER LOGIC =====================
function buildTasksNotesGrid(deliverable) {
  const tasksNotesWrap = el("div", { className: "tasks-notes-grid" });

  const tasksCol = el("div", { className: "tn-col tasks-col" });
  tasksCol.appendChild(
    el("div", { className: "tn-heading", textContent: "Tasks" })
  );
  const tasksBody = el("div", { className: "tn-body tasks-body" });
  if (deliverable.tasks && deliverable.tasks.length) {
    const renderTasks = (expanded) => {
      tasksBody.innerHTML = "";
      const tasksToShow = expanded
        ? deliverable.tasks.map((task, index) => ({ task, index }))
        : deliverable.tasks
          .slice(0, 2)
          .map((task, index) => ({ task, index }));
      tasksToShow.forEach(({ task, index }) => {
        const taskObj =
          typeof task === "string"
            ? { text: task, done: false, links: [] }
            : task;
        if (task !== taskObj) deliverable.tasks[index] = taskObj;
        const taskChip = el("div", {
          className: `task-chip ${taskObj.done ? "done" : ""}`,
          textContent: taskObj.text || "Task",
          role: "button",
          tabIndex: 0,
          "aria-pressed": String(!!taskObj.done),
        });
        const toggleTask = () => {
          taskObj.done = !taskObj.done;
          taskChip.classList.toggle("done", taskObj.done);
          taskChip.setAttribute("aria-pressed", String(!!taskObj.done));
          save();
        };
        taskChip.onclick = (e) => {
          e.stopPropagation();
          toggleTask();
        };
        taskChip.onkeydown = (e) => {
          if (e.key === "Enter" || e.key === " ") {
            e.preventDefault();
            toggleTask();
          }
        };
        tasksBody.appendChild(taskChip);
      });
      if (deliverable.tasks.length > 2) {
        const moreBtn = el("button", {
          className: "btn-more-tasks",
          textContent: expanded
            ? "Show Less"
            : `+${deliverable.tasks.length - 2} more`,
          onclick: (e) => {
            e.stopPropagation();
            renderTasks(!expanded);
          },
        });
        tasksBody.appendChild(moreBtn);
      }
    };
    renderTasks(false);
  } else {
    tasksBody.textContent = "--";
  }
  tasksCol.appendChild(tasksBody);

  const notesCol = el("div", { className: "tn-col notes-col" });
  notesCol.appendChild(
    el("div", { className: "tn-heading", textContent: "Notes" })
  );
  const notesBody = el("div", { className: "tn-body notes-body" });
  const notesText = (deliverable.notes || "").trim();
  if (notesText) {
    notesBody.appendChild(
      el("div", { className: "note-snippet", textContent: notesText })
    );
  } else {
    notesBody.textContent = "--";
  }
  notesCol.appendChild(notesBody);

  tasksNotesWrap.append(tasksCol, notesCol);
  return tasksNotesWrap;
}

// ===================== CARD-BASED DELIVERABLE RENDERING =====================

function getTaskCompletionStats(deliverable) {
  if (!deliverable.tasks || !deliverable.tasks.length) {
    return { total: 0, completed: 0, percentage: 0 };
  }
  const total = deliverable.tasks.length;
  const completed = deliverable.tasks.filter(t => {
    const taskObj = typeof t === "string" ? { done: false } : t;
    return taskObj.done;
  }).length;
  const percentage = Math.round((completed / total) * 100);
  return { total, completed, percentage };
}

function createCardHeader(deliverable, isPrimary, card, project) {
  const header = el("div", { className: "deliverable-card-header-new" });

  // Left section: title + due date
  const leftSection = el("div", { className: "deliverable-header-left" });

  const title = el("div", { className: "deliverable-card-title-new" });

  if (isPrimary) {
    const starIcon = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    starIcon.setAttribute("viewBox", "0 0 24 24");
    starIcon.setAttribute("fill", "currentColor");
    starIcon.innerHTML = `<path d="${STAR_ICON_PATH}"/>`;
    title.appendChild(starIcon);
  }

  const nameSpan = el("span", {
    className: "deliverable-card-title-name",
    textContent: deliverable.name || "Deliverable",
    title: "Double-click to edit"
  });

  // Double-click to edit deliverable name
  nameSpan.addEventListener("dblclick", (e) => {
    e.stopPropagation();

    // Create input for inline editing
    const input = el("input", {
      className: "deliverable-name-input",
      type: "text",
      value: deliverable.name || ""
    });

    // Replace span with input
    nameSpan.style.display = "none";
    title.insertBefore(input, nameSpan.nextSibling);
    input.focus();
    input.select();

    const finishEditing = async (shouldSave) => {
      if (shouldSave && input.value.trim()) {
        deliverable.name = input.value.trim();
        nameSpan.textContent = deliverable.name;
        await save();
      }
      input.remove();
      nameSpan.style.display = "";
    };

    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        finishEditing(true);
      } else if (e.key === "Escape") {
        e.preventDefault();
        finishEditing(false);
      }
    });

    input.addEventListener("blur", () => {
      finishEditing(true);
    });

    input.addEventListener("click", (e) => {
      e.stopPropagation();
    });
  });

  title.appendChild(nameSpan);

  leftSection.appendChild(title);

  // Due date badge - now on the left after title
  if (deliverable.due) {
    const ds = dueState(deliverable.due);
    const badgeClass = `deliverable-due-badge ${ds === "overdue" ? "overdue" : ds === "dueSoon" ? "due-soon" : "ok"}`;
    const badgeText = ds === "overdue" ? "OVERDUE" : humanDate(deliverable.due).replace(/\//g, "/").toUpperCase();
    const badge = el("div", {
      className: `${badgeClass} is-clickable`,
      textContent: badgeText
    });
    badge.setAttribute("role", "button");
    badge.setAttribute("tabindex", "0");
    badge.setAttribute("title", "Click to change due date");
    badge.setAttribute(
      "aria-label",
      ds === "overdue"
        ? "Overdue. Click to change due date."
        : `Due ${humanDate(deliverable.due)}. Click to change due date.`
    );
    const handleOpen = (e) => {
      e.stopPropagation();
      showCalendarForDeliverableBadge(badge, deliverable, project);
    };
    badge.addEventListener("click", handleOpen);
    badge.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        handleOpen(e);
      }
    });
    leftSection.appendChild(badge);
  }

  header.appendChild(leftSection);

  // Right section: expand/contract toggle
  const expandToggle = createExpandToggle(card);
  header.appendChild(expandToggle);

  return header;
}

function createExpandToggle(card) {
  const btn = el("button", {
    className: "deliverable-expand-toggle",
    title: "Expand/collapse details"
  });

  // Create expand icon (outward arrows)
  const expandSvg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  expandSvg.setAttribute("viewBox", "0 0 24 24");
  expandSvg.setAttribute("fill", "none");
  expandSvg.setAttribute("stroke", "currentColor");
  expandSvg.setAttribute("stroke-width", "2");
  expandSvg.setAttribute("class", "expand-icon");
  // Expand icon - arrows pointing outward
  expandSvg.innerHTML = '<polyline points="15 3 21 3 21 9"></polyline><polyline points="9 21 3 21 3 15"></polyline><line x1="21" y1="3" x2="14" y2="10"></line><line x1="3" y1="21" x2="10" y2="14"></line>';

  // Create contract icon (inward arrows)
  const contractSvg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  contractSvg.setAttribute("viewBox", "0 0 24 24");
  contractSvg.setAttribute("fill", "none");
  contractSvg.setAttribute("stroke", "currentColor");
  contractSvg.setAttribute("stroke-width", "2");
  contractSvg.setAttribute("class", "contract-icon");
  // Contract icon - arrows pointing inward
  contractSvg.innerHTML = '<polyline points="4 14 10 14 10 20"></polyline><polyline points="20 10 14 10 14 4"></polyline><line x1="14" y1="10" x2="21" y2="3"></line><line x1="3" y1="21" x2="10" y2="14"></line>';

  btn.append(expandSvg, contractSvg);

  btn.onclick = (e) => {
    e.stopPropagation();
    setDeliverableDetailsCollapsed(
      card,
      !card.classList.contains("details-collapsed")
    );
  };

  return btn;
}

function createProgressSection(deliverable) {
  const section = el("div", { className: "deliverable-progress-section" });
  const stats = getTaskCompletionStats(deliverable);

  // Progress bar
  const barContainer = el("div", { className: "deliverable-progress-bar-container" });
  const barFill = el("div", {
    className: "deliverable-progress-bar-fill",
    style: `width: ${stats.percentage}%`
  });
  barContainer.appendChild(barFill);

  // Progress text
  const progressClass = stats.percentage >= 50 ? "high" : stats.percentage >= 25 ? "medium" : "low";
  const progressText = el("div", { className: `deliverable-progress-text ${progressClass}` });

  const percentageSpan = el("span", {
    className: "percentage",
    textContent: `${stats.percentage}%`
  });
  const detailSpan = el("span", {
    className: "detail",
    textContent: stats.total > 0 ? ` complete (${stats.completed}/${stats.total} tasks)` : " (no tasks)"
  });

  progressText.append(percentageSpan, detailSpan);

  section.append(barContainer, progressText);
  return section;
}

function createStatusBadges(deliverable) {
  const container = el("div", { className: "deliverable-status-badges" });

  if (deliverable.statuses && deliverable.statuses.length) {
    deliverable.statuses.forEach(status => {
      const statusClass = status.toLowerCase().replace(/\s+/g, "");
      const badge = el("div", {
        className: `deliverable-status-badge ${statusClass}`,
        textContent: status
      });
      container.appendChild(badge);
    });
  }

  return container;
}

function createStatusDropdown(deliverable, project, card) {
  const availableStatuses = ["Waiting", "Working", "Pending Review", "Complete", "Delivered"];
  const dropdown = el("div", { className: "deliverable-status-dropdown" });

  // Trigger button - compact ellipsis icon
  const trigger = el("button", {
    className: "deliverable-status-trigger",
    title: "Status options",
    "aria-label": "Status options"
  });

  const dotsSvg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  dotsSvg.setAttribute("viewBox", "0 0 24 24");
  dotsSvg.setAttribute("fill", "currentColor");
  dotsSvg.innerHTML = '<circle cx="6" cy="12" r="1.6"></circle><circle cx="12" cy="12" r="1.6"></circle><circle cx="18" cy="12" r="1.6"></circle>';

  trigger.appendChild(dotsSvg);

  // Dropdown menu - status options only (no Show details)
  const menu = el("div", { className: "deliverable-status-menu" });

  availableStatuses.forEach(status => {
    const isActive = deliverable.statuses && deliverable.statuses.includes(status);
    const option = el("label", { className: "deliverable-status-option" });

    const checkbox = el("input", {
      type: "checkbox",
      checked: isActive
    });

    checkbox.addEventListener("change", async (e) => {
      e.stopPropagation();

      if (!deliverable.statuses) {
        deliverable.statuses = [];
      }

      if (checkbox.checked) {
        if (!deliverable.statuses.includes(status)) {
          deliverable.statuses.push(status);
        }
      } else {
        deliverable.statuses = deliverable.statuses.filter(s => s !== status);
      }

      deliverable.statusTags = deliverable.statuses
        .map((s) => LABEL_TO_KEY[s])
        .filter(Boolean);
      syncStatusArrays(deliverable);

      // Save changes
      await save();

      // Update the status badges display
      const card = dropdown.closest(".deliverable-card-new");
      const badgesContainer = card.querySelector(".deliverable-status-badges");
      if (badgesContainer) {
        badgesContainer.innerHTML = "";
        if (deliverable.statuses.length) {
          deliverable.statuses.forEach(s => {
            const statusClass = s.toLowerCase().replace(/\s+/g, "");
            const badge = el("div", {
              className: `deliverable-status-badge ${statusClass}`,
              textContent: s
            });
            badgesContainer.appendChild(badge);
          });
        }
      }
    });

    const label = el("span", { textContent: status });
    option.append(checkbox, label);
    menu.appendChild(option);
  });

  // Toggle dropdown
  trigger.addEventListener("click", (e) => {
    e.stopPropagation();
    const isOpen = menu.classList.toggle("open");
    trigger.classList.toggle("open", isOpen);
    card.classList.toggle("deliverable-menu-open", isOpen);

    // Close other open dropdowns
    document.querySelectorAll(".deliverable-status-menu.open").forEach(m => {
      if (m !== menu) {
        m.classList.remove("open");
        m.previousElementSibling.classList.remove("open");
        m.closest(".deliverable-card-new")?.classList.remove("deliverable-menu-open");
      }
    });
  });

  // Close dropdown when clicking outside
  document.addEventListener("click", (e) => {
    if (!dropdown.contains(e.target)) {
      menu.classList.remove("open");
      trigger.classList.remove("open");
      card.classList.remove("deliverable-menu-open");
    }
  });

  dropdown.append(trigger, menu);
  return dropdown;
}

function createNotesSection(deliverable, project) {
  const section = el("div", { className: "deliverable-notes-section" });

  // Simple header label (no toggle)
  const header = el("div", { className: "deliverable-notes-header-simple" });
  const titleText = el("span", {
    className: "deliverable-notes-label",
    textContent: "Notes"
  });
  header.appendChild(titleText);

  // Content (textarea) - always visible when details are expanded
  const textarea = el("textarea", {
    className: "deliverable-notes-textarea",
    placeholder: "Add notes about this deliverable...",
    value: deliverable.notes || ""
  });

  // Auto-save on change
  let saveTimeout;
  textarea.addEventListener("input", () => {
    deliverable.notes = textarea.value;

    // Debounce save
    clearTimeout(saveTimeout);
    saveTimeout = setTimeout(async () => {
      await save();
    }, 500);
  });

  section.append(header, textarea);
  return section;
}

function createTasksPreview(deliverable, card) {
  const container = el("div", { className: "deliverable-tasks-preview" });

  // Initialize tasks array if it doesn't exist
  if (!deliverable.tasks) {
    deliverable.tasks = [];
  }

  const stats = getTaskCompletionStats(deliverable);
  const heading = el("div", {
    className: "deliverable-tasks-preview-heading",
    textContent: deliverable.tasks.length > 0 ? `Tasks (${stats.completed}/${stats.total}):` : "Tasks:"
  });

  const list = el("div", { className: "deliverable-tasks-preview-list" });
  const maxVisible = 3;

  // Helper to update heading and progress bar
  const updateStatsDisplay = () => {
    const newStats = getTaskCompletionStats(deliverable);
    heading.textContent = deliverable.tasks.length > 0 ? `Tasks (${newStats.completed}/${newStats.total}):` : "Tasks:";

    const progressSection = card.querySelector(".deliverable-progress-section");
    if (progressSection) {
      const barFill = progressSection.querySelector(".deliverable-progress-bar-fill");
      const progressText = progressSection.querySelector(".deliverable-progress-text");
      const percentageSpan = progressText.querySelector(".percentage");
      const detailSpan = progressText.querySelector(".detail");

      barFill.style.width = `${newStats.percentage}%`;
      percentageSpan.textContent = `${newStats.percentage}%`;
      detailSpan.textContent = newStats.total > 0 ? ` complete (${newStats.completed}/${newStats.total} tasks)` : " (no tasks)";
      progressText.className = `deliverable-progress-text ${newStats.percentage >= 50 ? "high" : newStats.percentage >= 25 ? "medium" : "low"}`;
    }
  };

  const renderTaskList = (showAll) => {
    list.innerHTML = "";
    const tasksToShow = showAll ? deliverable.tasks : deliverable.tasks.slice(0, maxVisible);

    tasksToShow.forEach((task, index) => {
      const taskObj = typeof task === "string" ? { text: task, done: false } : task;
      const item = el("div", {
        className: `deliverable-task-item ${taskObj.done ? "done" : "undone"}`
      });

      // Checkmark or circle icon
      const icon = taskObj.done ? "✓" : "○";
      const iconSpan = el("span", { className: "task-icon", textContent: icon });
      const textSpan = el("span", { className: "task-text", textContent: taskObj.text || "Task" });

      // Delete button (visible on hover)
      const deleteBtn = el("button", {
        className: "task-delete-btn",
        title: "Remove task",
        textContent: "×"
      });

      deleteBtn.addEventListener("click", async (e) => {
        e.stopPropagation();

        // Find actual index in case we're showing limited view
        const actualIndex = showAll ? index : index;
        deliverable.tasks.splice(actualIndex, 1);

        await save();
        updateStatsDisplay();
        renderTaskList(showAll && deliverable.tasks.length > maxVisible);
      });

      item.append(iconSpan, textSpan, deleteBtn);

      // Make task clickable to toggle completion (but not on delete button)
      item.addEventListener("click", async (e) => {
        if (e.target === deleteBtn) return;
        e.stopPropagation();

        // Toggle done state
        taskObj.done = !taskObj.done;

        // Update the task in the array
        const actualIndex = showAll ? index : index;
        if (typeof deliverable.tasks[actualIndex] === "string") {
          deliverable.tasks[actualIndex] = { text: deliverable.tasks[actualIndex], done: taskObj.done };
        } else {
          deliverable.tasks[actualIndex].done = taskObj.done;
        }

        // Update UI
        item.classList.toggle("done", taskObj.done);
        item.classList.toggle("undone", !taskObj.done);
        iconSpan.textContent = taskObj.done ? "✓" : "○";

        await save();
        updateStatsDisplay();
      });

      list.appendChild(item);
    });

    // "+N more" button if there are more tasks and we're not showing all
    if (deliverable.tasks.length > maxVisible && !showAll) {
      const moreBtn = el("button", {
        className: "deliverable-expand-btn",
        textContent: `+${deliverable.tasks.length - maxVisible} more`,
        onclick: (e) => {
          e.stopPropagation();
          renderTaskList(true);
        }
      });
      list.appendChild(moreBtn);
    } else if (showAll && deliverable.tasks.length > maxVisible) {
      const lessBtn = el("button", {
        className: "deliverable-expand-btn",
        textContent: "Show less",
        onclick: (e) => {
          e.stopPropagation();
          renderTaskList(false);
        }
      });
      list.appendChild(lessBtn);
    }

    // Add new task input at the bottom
    const addTaskRow = el("div", { className: "task-add-row" });
    const bulletSpan = el("span", { className: "task-icon task-add-bullet", textContent: "○" });
    const taskInput = el("input", {
      className: "task-add-input",
      type: "text",
      placeholder: "Add a task..."
    });

    taskInput.addEventListener("keydown", async (e) => {
      if (e.key === "Enter" && taskInput.value.trim()) {
        e.preventDefault();
        e.stopPropagation();

        const newTask = { text: taskInput.value.trim(), done: false };
        deliverable.tasks.push(newTask);

        await save();
        updateStatsDisplay();

        // Re-render showing all if we were showing all, otherwise show limited
        const shouldShowAll = showAll || deliverable.tasks.length <= maxVisible;
        renderTaskList(shouldShowAll);

        // Focus the input after re-render for rapid task entry
        setTimeout(() => {
          const newTaskInput = list.querySelector(".task-add-input");
          if (newTaskInput) {
            newTaskInput.focus();
            // Select all text for easy replacement
            newTaskInput.select();
          }
        }, 0);
      }
    });

    taskInput.addEventListener("click", (e) => {
      e.stopPropagation();
    });

    addTaskRow.append(bulletSpan, taskInput);
    list.appendChild(addTaskRow);
  };

  renderTaskList(false);

  container.append(heading, list);
  return container;
}

async function refreshTimesheetsInfo() {
  const pathInput = document.getElementById("settings_timesheetsPath");
  const infoEl = document.getElementById("settings_timesheetsInfo");
  if (!pathInput || !infoEl) return;

  try {
    const res = await window.pywebview.api.get_timesheets_info();
    if (res && res.status !== "success") {
      throw new Error(res.message || "Failed to read timesheets info.");
    }

    pathInput.value = res.path || "";
    if (!res.exists) {
      infoEl.textContent = "File not found yet.";
    } else {
      const modified = res.modified
        ? new Date(res.modified).toLocaleString()
        : "unknown";
      infoEl.textContent = `Last modified: ${modified} • Size: ${res.size} bytes`;
    }
  } catch (e) {
    console.warn("Failed to refresh timesheets info:", e);
    infoEl.textContent = "Unable to read timesheets file info.";
  }
}

function setDeliverableDetailsCollapsed(card, isCollapsed) {
  card.classList.toggle("details-collapsed", isCollapsed);
}

function renderDeliverableCard(deliverable, isPrimary, project) {
  const card = el("div", {
    className: `deliverable-card-new ${isPrimary ? "is-primary" : ""} details-collapsed`
  });

  // Header: name + due badge + expand toggle (pass card for toggle)
  const header = createCardHeader(deliverable, isPrimary, card, project);

  // Progress: bar + percentage text
  const progress = createProgressSection(deliverable);

  // Status section: badges + dropdown inline
  const statusSection = el("div", { className: "deliverable-status-row" });
  const statusBadges = createStatusBadges(deliverable);
  const statusDropdown = createStatusDropdown(deliverable, project, card);
  statusSection.append(statusBadges, statusDropdown);

  // Tasks preview (2-3 tasks, now clickable)
  const tasksPreview = createTasksPreview(deliverable, card);

  // Notes section (always visible when expanded, no toggle)
  const notesSection = createNotesSection(deliverable, project);

  card.append(header, progress, statusSection, tasksPreview, notesSection);
  return card;
}

function normalizeProjectMatchValue(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function normalizeProjectId(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "")
    .trim();
}

function getNameTokens(value) {
  const normalized = normalizeProjectMatchValue(value);
  if (!normalized) return [];
  return normalized.split(" ").filter((token) => token.length > 2);
}

function nameSimilarityScore(a, b) {
  if (!a || !b) return 0;
  if (a === b) return 1;
  if (a.includes(b) || b.includes(a)) return 0.9;
  const aTokens = new Set(getNameTokens(a));
  const bTokens = new Set(getNameTokens(b));
  if (!aTokens.size || !bTokens.size) return 0;
  let intersect = 0;
  aTokens.forEach((token) => {
    if (bTokens.has(token)) intersect++;
  });
  const union = aTokens.size + bTokens.size - intersect;
  return union ? intersect / union : 0;
}

function areProjectsSimilar(a, b) {
  if (!a || !b) return false;
  const idA = normalizeProjectId(a.id);
  const idB = normalizeProjectId(b.id);
  if (idA && idB) {
    if (idA === idB) return true;
    if (
      idA.length >= 4 &&
      idB.length >= 4 &&
      (idA.includes(idB) || idB.includes(idA))
    )
      return true;
  }
  const nameA = normalizeProjectMatchValue(a.name);
  const nameB = normalizeProjectMatchValue(b.name);
  if (nameA && nameB) {
    const score = nameSimilarityScore(nameA, nameB);
    if (score >= 0.8) return true;
  }
  return false;
}

function findBestProjectMatch(aiProject) {
  if (!aiProject || !db.length) return null;
  const MATCH_THRESHOLD = 0.8;
  const aiId = normalizeProjectId(aiProject.id);
  const aiBaseKey = getProjectBaseKey(aiProject.path);
  const aiName = normalizeProjectMatchValue(aiProject.name);
  const aiNick = normalizeProjectMatchValue(aiProject.nick);
  let bestMatch = null;
  for (let i = 0; i < db.length; i++) {
    const existing = db[i];
    if (!existing) continue;
    let candidateScore = 0;
    let candidateMethod = "";
    const existingId = normalizeProjectId(existing.id);
    if (aiId && existingId) {
      if (aiId === existingId) {
        candidateScore = 1.0;
        candidateMethod = "id-exact";
      } else if (
        aiId.length >= 4 &&
        existingId.length >= 4 &&
        (aiId.includes(existingId) || existingId.includes(aiId))
      ) {
        candidateScore = 0.95;
        candidateMethod = "id-substring";
      }
    }
    if (candidateScore < 0.9 && aiBaseKey) {
      const existingBaseKey = getProjectBaseKey(existing.path);
      if (existingBaseKey && aiBaseKey === existingBaseKey) {
        candidateScore = 0.9;
        candidateMethod = "path";
      }
    }
    if (candidateScore < MATCH_THRESHOLD && aiName) {
      const existingName = normalizeProjectMatchValue(existing.name);
      if (existingName) {
        const nameScore = nameSimilarityScore(aiName, existingName);
        if (nameScore >= MATCH_THRESHOLD && nameScore > candidateScore) {
          candidateScore = nameScore;
          candidateMethod = "name";
        }
      }
    }
    if (candidateScore < MATCH_THRESHOLD) {
      const existingNick = normalizeProjectMatchValue(existing.nick);
      if (existingNick) {
        if (aiName) {
          const nickScore = nameSimilarityScore(aiName, existingNick) * 0.95;
          if (nickScore >= MATCH_THRESHOLD && nickScore > candidateScore) {
            candidateScore = nickScore;
            candidateMethod = "nick";
          }
        }
        if (aiNick) {
          const nickScore2 = nameSimilarityScore(aiNick, existingNick) * 0.95;
          if (nickScore2 >= MATCH_THRESHOLD && nickScore2 > candidateScore) {
            candidateScore = nickScore2;
            candidateMethod = "nick";
          }
        }
      }
    }
    if (candidateScore >= MATCH_THRESHOLD) {
      if (!bestMatch || candidateScore > bestMatch.score) {
        bestMatch = { index: i, score: candidateScore, method: candidateMethod };
      }
    }
  }
  return bestMatch;
}

function pickShortestPath(projects = []) {
  let shortest = "";
  projects.forEach((p) => {
    const path = String(p?.path || "").trim();
    if (!path) return;
    if (!shortest || path.length < shortest.length) shortest = path;
  });
  return shortest;
}

function scanAndMergeSimilarProjects() {
  if (!db.length) {
    toast("No projects to scan.");
    return;
  }
  if (!confirm("Scan all projects and merge similar matches?")) return;
  const total = db.length;
  const parent = Array.from({ length: total }, (_, i) => i);

  const find = (x) => {
    let root = x;
    while (parent[root] !== root) root = parent[root];
    while (parent[x] !== x) {
      const next = parent[x];
      parent[x] = root;
      x = next;
    }
    return root;
  };

  const union = (a, b) => {
    const ra = find(a);
    const rb = find(b);
    if (ra !== rb) parent[rb] = ra;
  };

  for (let i = 0; i < total; i++) {
    for (let j = i + 1; j < total; j++) {
      if (areProjectsSimilar(db[i], db[j])) union(i, j);
    }
  }

  const groups = new Map();
  for (let i = 0; i < total; i++) {
    const root = find(i);
    if (!groups.has(root)) groups.set(root, []);
    groups.get(root).push(i);
  }

  const mergeGroups = Array.from(groups.values()).filter((g) => g.length > 1);
  if (!mergeGroups.length) {
    toast("No similar projects found.");
    return;
  }

  mergeGroups.sort((a, b) => Math.max(...b) - Math.max(...a));
  let mergedCount = 0;
  let removedCount = 0;

  mergeGroups.forEach((indexes) => {
    const sorted = indexes.slice().sort((a, b) => a - b);
    const projects = sorted.map((idx) => normalizeProject(db[idx]));
    const base = normalizeProject(db[sorted[0]]);
    projects.slice(1).forEach((project) => mergeProjects(base, project));

    base.path = pickShortestPath(projects) || base.path || "";
    base.deliverables = projects
      .flatMap((project) =>
        getProjectDeliverables(project).map(normalizeDeliverable)
      )
      .map((deliverable) => ({
        ...deliverable,
        id: deliverable.id || createId("dlv"),
      }));
    base.deliverables.sort(compareDeliverablesByDueDesc);
    if (!base.deliverables.length) base.deliverables = [createDeliverable()];
    base.overviewDeliverableId = base.deliverables[0]?.id || "";
    autoSetPrimary(base);
    base.overviewSortDir = "desc";
    base.pinned = projects.some((project) => project?.pinned);

    db[sorted[0]] = base;
    for (let i = sorted.length - 1; i >= 1; i--) {
      db.splice(sorted[i], 1);
      removedCount++;
    }
    mergedCount++;
  });

  save();
  render();
  toast(`Merged ${mergedCount} group${mergedCount === 1 ? "" : "s"} (${removedCount} removed).`);
}

function sortProjectsByCurrent(items) {
  items.sort((a, b) => {
    const dir = currentSort.dir === "asc" ? 1 : -1;
    if (currentSort.key === "due") {
      const da = getProjectSortKey(a);
      const dbb = getProjectSortKey(b);
      if (!da && !dbb) return 0;
      if (!da) return 1;
      if (!dbb) return -1;
      return (da - dbb) * dir;
    }
    const valA = a[currentSort.key];
    const valB = b[currentSort.key];
    return (
      String(valA || "").localeCompare(String(valB || ""), undefined, {
        numeric: true,
      }) * dir
    );
  });
}

function sortProjectsByDueDesc(items) {
  items.sort((a, b) => {
    const da = getProjectSortKey(a);
    const dbb = getProjectSortKey(b);
    if (!da && !dbb) return 0;
    if (!da) return 1;
    if (!dbb) return -1;
    return dbb - da;
  });
}

function render() {
  const tbody = document.getElementById("tbody");
  const emptyState = document.getElementById("emptyState");
  tbody.innerHTML = "";

  const q = val("search").toLowerCase();

  let items = db.filter((p) => {
    if (q && !matches(q, p)) return false;
    const overviewDeliverables = getOverviewDeliverables(p);
    if (!overviewDeliverables.length) return false;
    const primaryDeliverable = getPrimaryDeliverable(p);
    if (!primaryDeliverable) return false;
    if (dueFilter !== "all") {
      if (!matchesDueFilter(primaryDeliverable, dueFilter)) return false;
    }
    if (statusFilter === "incomplete") {
      if (isFinished(primaryDeliverable)) return false;
    } else if (
      statusFilter !== "all" &&
      !hasStatus(primaryDeliverable, statusFilter)
    ) {
      return false;
    }
    return true;
  });

  const pinned = items.filter((p) => p?.pinned);
  const unpinned = items.filter((p) => !p?.pinned);
  sortProjectsByDueDesc(pinned);
  sortProjectsByCurrent(unpinned);
  items = pinned.concat(unpinned);

  updateSortHeaders();

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const dayOfWeek = today.getDay();
  const startOfWeek = new Date(today);
  startOfWeek.setDate(today.getDate() - dayOfWeek);
  const endOfWeek = new Date(startOfWeek);
  endOfWeek.setDate(startOfWeek.getDate() + 6);
  endOfWeek.setHours(23, 59, 59, 999);
  const startOfLastWeek = new Date(startOfWeek);
  startOfLastWeek.setDate(startOfWeek.getDate() - 7);
  const endOfLastWeek = new Date(endOfWeek);
  endOfLastWeek.setDate(endOfWeek.getDate() - 7);
  const currentYear = today.getFullYear();
  const lastYear = currentYear - 1;

  let dueThisWeek = 0,
    dueLastWeek = 0,
    upcoming = 0,
    completedThisYear = 0,
    completedLastYear = 0;
  let minDate = null,
    maxDate = null;

  getAllDeliverables().forEach(({ deliverable }) => {
    const d = parseDueStr(deliverable?.due);
    if (d) {
      if (d >= startOfWeek && d <= endOfWeek) dueThisWeek++;
      if (d >= startOfLastWeek && d <= endOfLastWeek) dueLastWeek++;
      if (d > endOfWeek) upcoming++;
      if (!minDate || d < minDate) minDate = d;
      if (!maxDate || d > maxDate) maxDate = d;
      if (isFinished(deliverable)) {
        if (d.getFullYear() === currentYear) completedThisYear++;
        if (d.getFullYear() === lastYear) completedLastYear++;
      }
    }
  });

  // Stats are now handled exclusively by the Statistics Modal (renderStats)
  // We do not update them here to avoid conflicts or incorrect "All Time" overwrites.

  emptyState.style.display = items.length ? "none" : "block";
  const rowTemplate = document.getElementById("project-row-template");

  items.forEach((p) => {
    const tr = rowTemplate.content.cloneNode(true).querySelector("tr");
    const idx = db.indexOf(p);
    const overviewDeliverables = getOverviewDeliverables(p);
    const primary = getPrimaryDeliverable(p);
    const totalDeliverables = getProjectDeliverables(p).length;
    const priorityId = primary?.id;

    const selectCell = tr.querySelector(".cell-select");
    const pinBtn = selectCell?.querySelector(".pin-btn");
    if (pinBtn) {
      const isPinned = !!p.pinned;
      pinBtn.classList.toggle("is-pinned", isPinned);
      pinBtn.setAttribute("aria-pressed", String(isPinned));
      pinBtn.setAttribute("aria-label", isPinned ? "Unpin project" : "Pin project");
      pinBtn.setAttribute("title", isPinned ? "Unpin project" : "Pin project");
      pinBtn.textContent = "";
      pinBtn.appendChild(createIcon(PIN_ICON_PATH, 14));
      pinBtn.onclick = async (e) => {
        e.stopPropagation();
        p.pinned = !p.pinned;
        await save();
        render();
      };
    }

    const idCell = tr.querySelector(".cell-id");
    const idBadge = idCell.querySelector(".id-badge") || idCell;
    idBadge.textContent = p.id || "--";

    const nameCell = tr.querySelector(".cell-name");
    if (p.path) {
      const link = el("button", {
        className: "path-link",
        textContent: p.name || "--",
        title: `Open: ${p.path}`,
      });
      link.onclick = async () => {
        try {
          await window.pywebview.api.open_path(convertPath(p.path));
          toast("Opening folder...");
        } catch (e) {
          toast("Failed to open path.");
        }
      };
      nameCell.appendChild(link);
    } else {
      nameCell.textContent = p.name || "--";
    }

    // Add account if available (from P:\ or M:\ drive path)
    const account = extractAccountFromPath(p.path);
    if (account) {
      nameCell.append(
        el("small", { className: "muted", textContent: ` (${account})` })
      );
    }

    // Add nickname if available
    if (p.nick) {
      nameCell.append(
        el("small", { className: "muted", textContent: ` (${p.nick})` })
      );
    }
    const projectNotes = (p.notes || "").trim();
    if (projectNotes) {
      nameCell.append(
        el("div", {
          className: "project-notes-snippet",
          textContent: projectNotes,
        })
      );
    }

    const deliverablesCell = tr.querySelector(".cell-deliverables");
    deliverablesCell.innerHTML = "";

    const visibleDeliverables = overviewDeliverables.filter(d => {
      if (hideNonPrimary && priorityId) return d.id === priorityId;
      return true;
    });

    if (visibleDeliverables.length) {
      const cardsContainer = el("div", { className: "deliverable-cards-container" });

      visibleDeliverables.forEach((deliverable) => {
        const isPrimary = deliverable.id === priorityId;
        const card = renderDeliverableCard(deliverable, isPrimary, p);
        cardsContainer.appendChild(card);
      });

      deliverablesCell.appendChild(cardsContainer);
    } else {
      deliverablesCell.textContent = "--";
    }

    const actionsCell = tr.querySelector(".cell-actions");
    const actionsStack = el("div", { className: "actions-stack" });
    actionsStack.append(
      createIconButton({
        className: "btn icon-only",
        title: "Edit",
        ariaLabel: "Edit project",
        path: PENCIL_ICON_PATH,
        onClick: () => openEdit(idx),
      }),
      createIconButton({
        className: "btn icon-only",
        title: "Duplicate",
        ariaLabel: "Duplicate project",
        path: COPY_ICON_PATH,
        onClick: () => duplicate(idx),
      }),
      createIconButton({
        className: "btn btn-danger icon-only",
        title: "Delete",
        ariaLabel: "Delete project",
        path: TRASH_ICON_PATH,
        onClick: () => removeProject(idx),
      })
    );
    actionsCell.append(actionsStack);
    tbody.appendChild(tr);
  });
}

function renderStatusToggles(p) {
  const wrap = el("div", { className: "status-group" });
  const mk = (cls, label) => {
    const b = el("button", {
      className: `st status-btn st-${cls}`,
      type: "button",
      textContent: label,
      title: label,
      "aria-pressed": String(hasStatus(p, label)),
    });
    b.onclick = async (e) => {
      e.stopPropagation();
      toggleStatus(p, label);
      await save();
      render();
    };
    return b;
  };
  [
    ["wait", "Waiting"],
    ["work", "Working"],
    ["pr", "Pending Review"],
    ["comp", "Complete"],
    ["del", "Delivered"],
  ].forEach(([cls, label]) => wrap.append(mk(cls, label)));
  return wrap;
}

function matches(q, p) {
  if (!q) return true;
  const str = (val) => (val || "").toLowerCase();
  if (
    str(p.id).includes(q) ||
    str(p.name).includes(q) ||
    str(p.nick).includes(q) ||
    str(p.notes).includes(q)
  )
    return true;
  return getProjectDeliverables(p).some((d) => {
    if (
      str(d.name).includes(q) ||
      str(d.notes).includes(q) ||
      str(d.due).includes(q)
    )
      return true;
    if ((d.tasks || []).some((t) => str(t.text).includes(q))) return true;
    return (d.statuses || []).some((s) => str(s).includes(q));
  });
}

function updateSortHeaders() {
  document.querySelectorAll("th[data-sort]").forEach((th) => {
    th.classList.remove("sort-asc", "sort-desc");
    if (th.dataset.sort === currentSort.key)
      th.classList.add(`sort-${currentSort.dir}`);
  });
}

// ===================== CRUD OPERATIONS =====================
function createBlankProject() {
  const deliverable = createDeliverable();
  return {
    id: "",
    name: "",
    nick: "",
    path: "",
    notes: "",
    refs: [],
    deliverables: [deliverable],
    overviewDeliverableId: deliverable.id,
  };
}

function openEdit(i) {
  editIndex = i;
  const p = db[i];
  document.getElementById("dlgTitle").textContent = `Edit Project — ${p.id || "Untitled"
    }`;
  document.getElementById("btnSaveProject").textContent = "Save Changes";
  fillForm(p);
  document.getElementById("editDlg").showModal();
}
function openNew() {
  editIndex = -1;
  document.getElementById("dlgTitle").textContent = "New Project";
  document.getElementById("btnSaveProject").textContent = "Create Project";
  fillForm(createBlankProject());
  document.getElementById("editDlg").showModal();
}
function removeProject(i) {
  if (!confirm("Delete this project?")) return;
  db.splice(i, 1);
  save();
  render();
}
function duplicate(i) {
  const original = db[i];
  const deliverable = createDeliverable();
  const newProjectData = {
    id: original?.id || "",
    name: original?.name || "",
    path: original?.path || "",
    nick: original?.nick || "",
    notes: "",
    refs: [],
    deliverables: [deliverable],
    overviewDeliverableId: deliverable.id,
  };
  editIndex = -1;
  document.getElementById("dlgTitle").textContent = "Duplicate Project";
  document.getElementById("btnSaveProject").textContent = "Create Duplicate";
  fillForm(newProjectData);
  document.getElementById("editDlg").showModal();
}
function onSaveProject() {
  // Validate all due dates first
  if (!validateAllDueDates()) {
    // Focus first invalid input
    const firstInvalid = document.querySelector('.d-due.input-error');
    if (firstInvalid) {
      firstInvalid.focus();
      // Scroll to the invalid input
      firstInvalid.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
    toast("Please fix invalid date formats before saving.", "error");
    return;
  }

  const data = readForm();
  autoSetPrimary(data);
  _aiMatchSnapshot = null;
  if (editIndex >= 0) {
    db[editIndex] = data;
    toast("Project updated successfully.");
  } else {
    db.push(data);
    toast("Project created.");
  }
  save();
  render();
  closeDlg("editDlg");
}

// ===================== FORM HANDLING =====================
function fillForm(project) {
  const p = normalizeProject(project || createBlankProject());
  document.getElementById("f_id").value = p.id || "";
  document.getElementById("f_name").value = p.name || "";
  document.getElementById("f_nick").value = p.nick || "";
  document.getElementById("f_notes").value = p.notes || "";
  document.getElementById("f_path").value = p.path || "";

  const deliverableList = document.getElementById("deliverableList");
  deliverableList.innerHTML = "";
  const sortedDeliverables = p.deliverables.slice();
  sortDeliverablesByPrimaryThenDueDesc(
    sortedDeliverables,
    p.overviewDeliverableId
  );
  sortedDeliverables.forEach((deliverable) =>
    addDeliverableCard(deliverable, p.overviewDeliverableId)
  );
  if (!deliverableList.children.length) {
    addDeliverableCard(createDeliverable(), p.overviewDeliverableId);
  }
  if (!deliverableList.querySelector(".d-primary:checked")) {
    const firstPrimary = deliverableList.querySelector(".d-primary");
    if (firstPrimary) firstPrimary.checked = true;
  }

  document.getElementById("refList").innerHTML = "";
  (p.refs || []).forEach(addRefRowFrom);
}

function addDeliverableCard(deliverable, primaryId, options = {}) {
  const list = document.getElementById("deliverableList");
  const template = document.getElementById("deliverable-card-template");
  if (!list || !template) return;
  const { insertAtTop = false } = options;

  const card = template.content
    .cloneNode(true)
    .querySelector(".deliverable-card");
  const deliverableId = deliverable.id || createId("dlv");
  card.dataset.deliverableId = deliverableId;

  card.querySelector(".d-name").value = deliverable.name || "";
  card.querySelector(".d-due").value = deliverable.due || "";
  card.querySelector(".d-notes").value = deliverable.notes || "";

  const primaryInput = card.querySelector(".d-primary");
  primaryInput.name = "primaryDeliverable";
  if (primaryId && deliverableId === primaryId) {
    primaryInput.checked = true;
  }

  // No secondary action needed on primary change anymore

  const taskList = card.querySelector(".deliverable-task-list");
  (deliverable.tasks || []).map(normalizeTask).forEach((task) => {
    addTaskRowFrom(taskList, task);
  });

  card.querySelector(".d-add-task").onclick = () =>
    addTaskRowFrom(taskList, {});
  card.querySelector(".btn-remove").onclick = () => card.remove();

  const statusContainer = card.querySelector(".deliverable-status");
  const markTasksDone = () => {
    taskList
      .querySelectorAll(".t-done")
      .forEach((cb) => (cb.checked = true));
  };
  const picker = buildStatusPicker(deliverable.statuses || [], (label, pressed) => {
    if (pressed && (label === "Complete" || label === "Delivered")) {
      markTasksDone();
    }
  });
  statusContainer.appendChild(picker);

  // Wire up calendar icon button and date validation
  const calendarBtn = card.querySelector('.calendar-icon-btn');
  const dueInput = card.querySelector('.d-due');

  if (calendarBtn && dueInput) {
    calendarBtn.onclick = (e) => {
      e.preventDefault();
      e.stopPropagation();
      showCalendarForInput(dueInput);
    };

    // Add auto-formatting and validation on blur
    dueInput.addEventListener('blur', () => {
      autoFormatDateInput(dueInput);
      validateDueDateInput(dueInput);
    });

    // Add validation on change
    dueInput.addEventListener('change', () => {
      autoFormatDateInput(dueInput);
      validateDueDateInput(dueInput);
    });
  }

  if (insertAtTop && list.firstChild) {
    list.prepend(card);
  } else {
    list.appendChild(card);
  }

  const hasPrimary = list.querySelector(".d-primary:checked");
  if (!hasPrimary && !primaryId) primaryInput.checked = true;

  return card;
}

function getMostRecentDeliverableCard() {
  const list = document.getElementById("deliverableList");
  if (!list) return null;
  const primaryInput = list.querySelector(".d-primary:checked");
  if (primaryInput) return primaryInput.closest(".deliverable-card");
  const cards = Array.from(list.querySelectorAll(".deliverable-card"));
  if (!cards.length) return null;
  let latestCard = null;
  let latestDue = null;
  cards.forEach((card) => {
    const due = parseDueStr(card.querySelector(".d-due")?.value.trim());
    if (!due) return;
    if (!latestDue || due > latestDue) {
      latestDue = due;
      latestCard = card;
    }
  });
  return latestCard || cards[0];
}

function isDeliverableCardComplete(card) {
  if (!card) return false;
  const statuses = readStatusPickerFrom(
    card.querySelector(".deliverable-status")
  );
  return statuses.includes("Complete");
}

function setPrimaryDeliverableCard(card) {
  const list = document.getElementById("deliverableList");
  if (!list || !card) return;
  list.querySelectorAll(".d-primary").forEach((input) => {
    input.checked = false;
  });
  const primaryInput = card.querySelector(".d-primary");
  if (primaryInput) primaryInput.checked = true;
}

function buildStatusPicker(selected = [], onToggle) {
  const wrap = el("div", { className: "status-picker" });
  const mk = (cls, label) => {
    const b = el("button", {
      className: `st status-btn st-${cls}`,
      type: "button",
      textContent: label,
      title: label,
      "data-status": label,
      "aria-pressed": String(selected.includes(label)),
    });
    b.onclick = (e) => {
      e.preventDefault();
      const next = b.getAttribute("aria-pressed") !== "true";
      b.setAttribute("aria-pressed", String(next));
      if (onToggle) onToggle(label, next, b);
    };
    return b;
  };
  [
    ["wait", "Waiting"],
    ["work", "Working"],
    ["pr", "Pending Review"],
    ["comp", "Complete"],
    ["del", "Delivered"],
  ].forEach(([cls, label]) => wrap.append(mk(cls, label)));
  return wrap;
}

function readStatusPickerFrom(container) {
  if (!container) return [];
  return Array.from(container.querySelectorAll('.st[aria-pressed="true"]')).map(
    (b) => b.dataset.status
  );
}

function readForm() {
  const baseSchedule =
    editIndex >= 0 && db[editIndex] ? db[editIndex].lightingSchedule : null;
  const existingProject = editIndex >= 0 && db[editIndex] ? db[editIndex] : null;
  const lightingSchedule = normalizeLightingSchedule(
    baseSchedule ?? createDefaultLightingSchedule()
  );
  const out = {
    id: val("f_id"),
    name: val("f_name"),
    nick: val("f_nick"),
    notes: val("f_notes"),
    path: val("f_path"),
    refs: [],
    deliverables: [],
    overviewDeliverableId: "",
    pinned: !!existingProject?.pinned,
    lightingSchedule,
  };

  document.querySelectorAll("#deliverableList .deliverable-card").forEach((card) => {
    const deliverableId = card.dataset.deliverableId || createId("dlv");
    const name = card.querySelector(".d-name").value.trim();
    const due = card.querySelector(".d-due").value.trim();
    const notes = card.querySelector(".d-notes").value;
    const statuses = readStatusPickerFrom(
      card.querySelector(".deliverable-status")
    );

    const tasks = [];
    card.querySelectorAll(".task-row").forEach((row) => {
      const text = row.querySelector(".t-text").value.trim();
      if (!text) return;
      const done = row.querySelector(".t-done").checked;
      const links = [
        row.querySelector(".t-link").value.trim(),
        row.querySelector(".t-link2").value.trim(),
      ]
        .filter(Boolean)
        .map(normalizeLink);
      tasks.push({ text, done, links });
    });

    const deliverable = normalizeDeliverable({
      id: deliverableId,
      name,
      due,
      notes,
      tasks,
      statuses,
    });

    out.deliverables.push(deliverable);
    if (card.querySelector(".d-primary").checked)
      out.overviewDeliverableId = deliverable.id;
  });

  if (!out.deliverables.length) {
    const deliverable = createDeliverable();
    out.deliverables = [deliverable];
    out.overviewDeliverableId = deliverable.id;
  }
  if (!out.overviewDeliverableId) {
    out.overviewDeliverableId = out.deliverables[0].id;
  }

  document.querySelectorAll("#refList .ref-row").forEach((row) => {
    const label = row.querySelector(".r-label").value.trim();
    const raw = row.querySelector(".r-url").value.trim();
    if (!raw) return;
    const link = normalizeLink(raw);
    if (label) link.label = label;
    out.refs.push(link);
  });
  return out;
}

function addTaskRowFrom(container, t = {}) {
  const template = document.getElementById("task-row-template");
  const row = template.content.cloneNode(true).querySelector(".task-row");
  const textInput = row.querySelector(".t-text");

  textInput.value = t.text || "";
  row.querySelector(".t-done").checked = !!t.done;
  row.querySelector(".t-link").value = t.links?.[0]?.raw || "";
  row.querySelector(".t-link2").value = t.links?.[1]?.raw || "";
  row.querySelector(".btn-remove").onclick = () => row.remove();

  // Add Enter key handler for rapid task entry
  textInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && textInput.value.trim()) {
      e.preventDefault();
      // Add new task row below and focus it
      addTaskRowFrom(container, {});
      // Focus the new row's text input
      setTimeout(() => {
        const newRow = row.nextElementSibling;
        const newInput = newRow?.querySelector(".t-text");
        if (newInput) newInput.focus();
      }, 0);
    }
  });

  if (container) container.append(row);
}

function addRefRowFrom(L = {}) {
  const template = document.getElementById("ref-row-template");
  const row = template.content.cloneNode(true).querySelector(".ref-row");
  row.querySelector(".r-label").value = L.label || "";
  row.querySelector(".r-url").value = L.raw || L.url || "";
  row.querySelector(".btn-remove").onclick = () => row.remove();
  document.getElementById("refList").append(row);
}

window.addDeliverable = () => {
  const mostRecentCard = getMostRecentDeliverableCard();
  const shouldSetPrimary =
    mostRecentCard && isDeliverableCardComplete(mostRecentCard);
  const newCard = addDeliverableCard(createDeliverable(), null, {
    insertAtTop: true,
  });
  if (shouldSetPrimary && newCard) setPrimaryDeliverableCard(newCard);
};
window.addRefRow = () => addRefRowFrom({});
window.closeDlg = closeDlg;

// ===================== CSV & IMPORT/EXPORT =====================
function parseCSV(text) {
  const rows = [];
  let i = 0,
    field = "",
    row = [],
    inQ = false;
  function pushField() {
    row.push(field);
    field = "";
  }
  function pushRow() {
    rows.push(row);
    row = [];
  }
  while (i < text.length) {
    const ch = text[i];
    if (inQ) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        }
        inQ = false;
        i++;
        continue;
      }
      field += ch;
      i++;
      continue;
    } else {
      if (ch === '"') {
        inQ = true;
        i++;
        continue;
      }
      if (ch === ",") {
        pushField();
        i++;
        continue;
      }
      if (ch === "\n") {
        pushField();
        pushRow();
        i++;
        continue;
      }
      if (ch === "\r") {
        if (text[i + 1] === "\n") {
          i += 2;
          pushField();
          pushRow();
        } else {
          i++;
          pushField();
          pushRow();
        }
        continue;
      }
      field += ch;
      i++;
    }
  }
  if (field !== "" || row.length) {
    pushField();
    pushRow();
  }
  return rows;
}
function importRows(rows, hasHeader = true) {
  if (rows.length && !hasHeader) {
    const joined = rows[0].map((s) => (s || "").toUpperCase()).join(" | ");
    if (joined.includes("PROJECT NAME") || joined.includes("DUE"))
      hasHeader = true;
  }
  if (hasHeader) rows = rows.slice(1);
  let added = 0;
  const incoming = [];

  for (const r of rows) {
    if (!r.length) continue;
    const [id, name, nick, notes, due, statusCell, tasksStr, path, ...refs] = r;
    if (!(id || name || tasksStr || refs.some(Boolean))) continue;

    const statuses = [];
    const parts = String(statusCell || "")
      .split(/[,/|;]+/)
      .map((s) => s.trim())
      .filter(Boolean);
    for (const s of parts) {
      const c = canonStatus(s);
      if (c && !statuses.includes(c)) statuses.push(c);
    }

    const tparts = (tasksStr || "")
      .replace(/\r/g, "\n")
      .split(/\n|;|\u2022|\r/)
      .map((s) => s.trim())
      .filter(Boolean);
    const tasks = tparts.map((t) => ({ text: t, done: false, links: [] }));

    const deliverableName =
      extractDeliverableName(tasksStr) ||
      extractDeliverableName(notes) ||
      extractDeliverableName(path) ||
      extractDeliverableName(name) ||
      "Deliverable";

    const deliverable = createDeliverable({
      name: deliverableName,
      due: (due || "").trim(),
      notes: (notes || "").trim(),
      tasks,
      statuses,
    });

    const refsList = [];
    for (const cell of refs) {
      const s = (cell || "").trim();
      if (!s) continue;
      refsList.push(normalizeLink(s));
    }

    incoming.push({
      id: String(id || "").trim(),
      name: (name || "").trim(),
      nick: (nick || "").trim(),
      notes: "",
      path: (path || "").trim(),
      refs: refsList,
      deliverables: [deliverable],
      overviewDeliverableId: deliverable.id,
      lightingSchedule: createDefaultLightingSchedule(),
    });
    added++;
  }

  if (incoming.length) {
    const merged = migrateProjects([...db, ...incoming]);
    db = merged.data;
  }
  save();
  render();
  toast(`Imported ${added} rows`);
}

// ===================== NOTES SYSTEM =====================
const debouncedSaveNotes = debounce(saveNotes, 500);

function renderNoteTabs() {
  const container = document.getElementById("notesTabsContainer");
  container.innerHTML = "";
  noteTabs.forEach((tabName) => {
    const btn = el("button", {
      className: `inner-tab-btn ${tabName === activeNoteTab ? "active" : ""}`,
      textContent: tabName,
      onclick: () => {
        activeNoteTab = tabName;
        renderNoteTabs();
        renderNoteSearchResults();
      },
    });
    const delIcon = el("span", {
      className: "tab-delete-icon",
      textContent: "🗑️",
      title: "Delete Page",
      onclick: (e) => {
        e.stopPropagation();
        if (confirm(`Permanently delete page "${tabName}"?`)) {
          const idx = noteTabs.indexOf(tabName);
          if (idx > -1) {
            noteTabs.splice(idx, 1);
            delete notesDb[tabName];
            if (activeNoteTab === tabName)
              activeNoteTab =
                noteTabs.length > 0 ? noteTabs[Math.max(0, idx - 1)] : null;
            saveNotes();
            renderNoteTabs();
          }
        }
      },
    });
    btn.appendChild(delIcon);
    container.appendChild(btn);
  });
  const addBtn = el("button", {
    className: "add-tab-btn",
    textContent: "+",
    title: "Add New Page",
    onclick: () => {
      const name = prompt("Enter name for new page:");
      if (name && name.trim()) {
        if (!noteTabs.includes(name.trim())) {
          noteTabs.push(name.trim());
          activeNoteTab = name.trim();
          saveNotes();
          renderNoteTabs();
        } else toast("Page name already exists.");
      }
    },
  });
  container.appendChild(addBtn);
  updateActiveNoteTextarea();
}

function updateActiveNoteTextarea() {
  const textarea = document.getElementById("notesTextarea");
  if (!activeNoteTab) {
    textarea.value = "";
    textarea.placeholder = "Create a page to begin.";
    textarea.disabled = true;
    return;
  }
  textarea.disabled = false;
  textarea.placeholder = `Enter notes for ${activeNoteTab}...`;
  textarea.value = notesDb[activeNoteTab] || "";
}

function handleNoteInput(e) {
  if (!activeNoteTab) return;
  notesDb[activeNoteTab] = e.target.value;
  debouncedSaveNotes();
}

function renderNoteSearchResults() {
  const query = val("notesSearch").toLowerCase();
  const resultsContainer = document.getElementById("notesSearchResults");
  resultsContainer.innerHTML = "";

  if (!query || !activeNoteTab) return;

  const queryWords = query.split(" ").filter((w) => w);
  if (!queryWords.length) return;

  const content = notesDb[activeNoteTab];
  if (!content) return;

  const notes = content.split(/\n\s*\n/).filter((note) => note.trim() !== "");

  notes.forEach((noteText) => {
    const lowerNoteText = noteText.toLowerCase();

    if (queryWords.every((word) => lowerNoteText.includes(word))) {
      const item = el("div", {
        className: "note-result-item",
        style: "position: relative;",
      });
      const contentDiv = el("div", { className: "note-result-content" });
      contentDiv.append(
        el("div", { className: "snippet", textContent: noteText })
      );

      const copyIcon = el("button", {
        className: "note-action-icon copy-icon",
        textContent: "📋",
        title: "Copy",
      });
      copyIcon.onclick = () => {
        navigator.clipboard.writeText(noteText).then(() => toast("Copied!"));
      };

      const editIcon = el("button", {
        className: "note-action-icon edit-icon",
        textContent: "✏️",
        title: "Edit",
      });

      editIcon.onclick = () => {
        // 1. Scroll the main Page View to the top immediately
        window.scrollTo({ top: 0, behavior: "smooth" });

        const textarea = document.getElementById("notesTextarea");
        const currentVal = textarea.value;
        const start = currentVal.indexOf(noteText);

        if (start !== -1) {
          const end = start + noteText.length;

          // 2. Focus and Select Text
          textarea.blur();
          textarea.setSelectionRange(start, end);
          textarea.focus();

          // 3. Calculate Exact Position using Mirror Div (Precision Scroll)
          const mirror = document.createElement("div");
          const style = window.getComputedStyle(textarea);

          mirror.style.visibility = "hidden";
          mirror.style.position = "absolute";
          mirror.style.top = "-9999px";
          mirror.style.whiteSpace = "pre-wrap";
          mirror.style.wordWrap = "break-word";

          [
            "fontFamily",
            "fontSize",
            "fontWeight",
            "lineHeight",
            "letterSpacing",
            "padding",
          ].forEach((p) => {
            mirror.style[p] = style[p];
          });

          mirror.style.boxSizing = "border-box";
          mirror.style.width = `${textarea.clientWidth}px`;
          mirror.style.border = "none";

          mirror.textContent = currentVal.substring(0, start);

          document.body.appendChild(mirror);
          const targetY = mirror.clientHeight;
          document.body.removeChild(mirror);

          // 4. Scroll the Textarea internally to the specific line
          textarea.scrollTop = Math.max(
            0,
            targetY - textarea.clientHeight * 0.3
          );
        } else {
          toast("Note content changed. Please refresh search.");
        }
      };

      item.append(contentDiv, copyIcon, editIcon);
      resultsContainer.append(item);
    }
  });
}

// ===================== BUNDLE / PLUGIN MANAGER =====================

// Fetch description from specific GitHub RAW url based on bundle name
async function fetchDescriptionForBundle(bundleName) {
  // Remove prefix/suffix to get core name. e.g., "ElectricalCommands.CleanCADCommands.bundle" -> "CleanCADCommands"
  const coreName = bundleName
    .replace("ElectricalCommands.", "")
    .replace(".bundle", "");

  // Check Cache
  if (DESCRIPTION_CACHE[coreName]) return DESCRIPTION_CACHE[coreName];

  // Construct RAW URL. Pattern: AutoCADCommands/<CoreName>/<CoreName>_descriptions.json
  const url = `https://raw.githubusercontent.com/jacobhusband/ElectricalCommands/main/AutoCADCommands/${coreName}/${coreName}_descriptions.json`;

  try {
    const res = await fetch(url);
    if (!res.ok) throw new Error("Not found");
    const json = await res.json();
    DESCRIPTION_CACHE[coreName] = json; // Cache it
    return json;
  } catch (e) {
    console.warn(`Could not fetch description for ${coreName}`, e);
    return null;
  }
}

function computeBundleKey(bundles) {
  if (!Array.isArray(bundles)) return "";
  return bundles
    .map((bundle) => {
      const name = bundle.name || "";
      const state = bundle.state || "";
      const local = bundle.local_version || "";
      const remote = bundle.remote_version || "";
      const bundleName = bundle.bundle_name || "";
      return `${name}|${state}|${local}|${remote}|${bundleName}`;
    })
    .sort()
    .join(";;");
}

async function fetchBundleStatuses({ silent = false } = {}) {
  if (bundlesLoadingPromise) return bundlesLoadingPromise;
  bundlesLoadingPromise = (async () => {
    const response = await window.pywebview.api.get_bundle_statuses();
    if (response.status !== "success") throw new Error(response.message);
    const data = Array.isArray(response.data) ? response.data : [];
    bundlesCache = data;
    bundlesCacheKey = computeBundleKey(data);
    bundlesLastCheckAt = Date.now();
    return data;
  })();

  try {
    return await bundlesLoadingPromise;
  } catch (e) {
    if (!silent) throw e;
    return null;
  } finally {
    bundlesLoadingPromise = null;
  }
}

function prefetchBundleDescriptions(bundles) {
  if (!Array.isArray(bundles) || bundles.length === 0) return;
  bundles.forEach((bundle) => {
    fetchDescriptionForBundle(bundle.name);
  });
}

async function renderBundles(bundles) {
  const container = document.getElementById("commands-container");
  if (!container) return;

  container.innerHTML = "";
  if (!Array.isArray(bundles) || bundles.length === 0) {
    container.textContent = "No command bundles found.";
    return;
  }

  const tagJobs = [];

  // Process each bundle
  for (const bundle of bundles) {
    // Normalize name
    const coreName = bundle.name
      .replace("ElectricalCommands.", "")
      .replace(".bundle", "");

    const card = el("div", { className: "release-card" });
    let statusClass, statusTitle, btnText, btnClass;

    if (bundle.state === "installed") {
      statusClass = "installed";
      statusTitle = `Installed (v${bundle.local_version})`;
      btnText = "Uninstall";
      btnClass = "btn-danger";
    } else if (bundle.state === "update_available") {
      statusClass = "update-available";
      statusTitle = `Update Available (v${bundle.remote_version})`;
      btnText = "Update";
      btnClass = "btn-accent";
    } else {
      statusClass = "not-installed";
      statusTitle = "Not Installed";
      btnText = "Install";
      btnClass = "btn-primary";
    }

    const header = el("div", { className: "release-card-header" }, [
      el("div", { className: "release-card-title" }, [
        el("div", {
          className: `bundle-status ${statusClass}`,
          title: statusTitle,
        }),
        el("span", { textContent: coreName }),
      ]),
    ]);

    const body = el("div", { className: "release-card-body" });
    const tags = el("div", { className: "command-tags" });
    body.append(tags);

    const footer = el("div", { className: "release-card-footer" });
    const btn = el("button", {
      className: `btn ${btnClass}`,
      textContent: btnText,
    });
    btn.dataset.bundleName = bundle.bundle_name;
    btn.dataset.actionType = btnText;

    // Note: We rely on 'bundle.asset' from the backend for installation URL.
    // If backend doesn't provide it, we assume standard release naming.
    if (bundle.state !== "installed" && bundle.asset) {
      btn.dataset.asset = JSON.stringify(bundle.asset);
    }

    footer.append(btn);
    card.append(header, body, footer);
    container.append(card);

    const descriptionPromise = fetchDescriptionForBundle(bundle.name)
      .then((description) => {
        if (description && description.commands) {
          Object.keys(description.commands).forEach((cmd) => {
            const link = description.links?.[cmd];
            const title =
              description.commands[cmd] || (link ? "Open documentation" : "");
            const tagEl = link
              ? el("button", {
                className: "command-tag command-link",
                textContent: cmd,
                title,
                onclick: () => openExternalUrl(link),
              })
              : el("span", { className: "command-tag", textContent: cmd, title });
            tags.append(tagEl);
          });
        }
      })
      .catch(() => { });
    tagJobs.push(descriptionPromise);
  }

  Promise.allSettled(tagJobs).catch(() => { });
}

async function ensureBundlesRendered({ force = false } = {}) {
  const container = document.getElementById("commands-container");
  if (!container) return;

  const shouldFetch = force || !bundlesCache;
  if (shouldFetch) {
    container.innerHTML = '<div class="spinner">Loading...</div>';
    try {
      await fetchBundleStatuses();
    } catch (e) {
      container.innerHTML = `<div class="error-message">Error: ${e.message}</div>`;
      return;
    }
  } else if (!force && bundlesLastRenderKey === bundlesCacheKey && !bundlesNeedsRender) {
    return;
  }

  await renderBundles(bundlesCache || []);
  bundlesLastRenderKey = bundlesCacheKey;
  bundlesNeedsRender = false;
}

function prefetchBundles() {
  if (bundlesPrefetchStarted) return;
  bundlesPrefetchStarted = true;
  setTimeout(() => {
    fetchBundleStatuses({ silent: true })
      .then((data) => {
        if (data) prefetchBundleDescriptions(data);
      })
      .catch(() => { });
  }, 0);
}

function setUpdateCheckIndicator(visible, text = "Checking for updates...") {
  const indicator = document.getElementById("updateCheckIndicator");
  if (!indicator) return;
  if (visible) {
    indicator.textContent = text;
    indicator.hidden = false;
    indicator.classList.add("visible");
  } else {
    indicator.classList.remove("visible");
    indicator.hidden = true;
  }
}

function isPluginsTabActive() {
  const panel = document.getElementById("plugins-panel");
  return !!panel && !panel.hidden;
}

async function checkBundlesForUpdates({ showIndicator = true } = {}) {
  if (bundlesUpdateCheckInFlight) return;
  const now = Date.now();
  if (now - bundlesLastCheckAt < 15000) return;

  bundlesUpdateCheckInFlight = true;
  if (showIndicator) setUpdateCheckIndicator(true);

  const previousKey = bundlesCacheKey;
  try {
    const data = await fetchBundleStatuses({ silent: true });
    if (!data) return;
    prefetchBundleDescriptions(data);
    if (bundlesCacheKey !== previousKey) {
      bundlesNeedsRender = true;
      if (isPluginsTabActive()) {
        await ensureBundlesRendered({ force: true });
      }
    }
  } finally {
    bundlesUpdateCheckInFlight = false;
    bundlesLastCheckAt = Date.now();
    if (showIndicator) setUpdateCheckIndicator(false);
  }
}

async function handleBundleAction(e) {
  const button = e.target.closest("[data-action-type]");
  if (!button) return;

  const actionType = button.dataset.actionType;
  button.disabled = true;
  button.textContent = "Processing...";

  try {
    let response;
    if (actionType === "Install" || actionType === "Update") {
      if (!button.dataset.asset) throw new Error("Installation data missing.");
      const asset = JSON.parse(button.dataset.asset);
      response = await window.pywebview.api.install_single_bundle(asset);
    } else {
      response = await window.pywebview.api.uninstall_bundle(
        button.dataset.bundleName
      );
    }
    if (response.status !== "success") throw new Error(response.message);
    toast(`${actionType} successful.`);
  } catch (err) {
    toast(`⚠️ ${err.message}`, 5000);
  } finally {
    await ensureBundlesRendered({ force: true });

  }
}

// ===================== TOOLS & SCRIPTS =====================
window.updateToolStatus = function (toolId, message) {
  const card = document.getElementById(toolId);
  if (!card) return;
  const statusEl = card.querySelector(".tool-card-status");
  const abortBtn = document.getElementById("abortBtn");
  statusEl.classList.remove("error");
  if (toolId === "toolCleanDwgs" && abortBtn) {
    if (message && message !== "DONE" && !message.startsWith("ERROR:")) {
      abortBtn.style.display = "inline-block";
      abortBtn.disabled = false;
      abortBtn.onclick = async () => {
        await window.pywebview.api.abort_clean_dwgs();
        toast("Aborting...");
        abortBtn.disabled = true;
      };
    } else {
      abortBtn.style.display = "none";
    }
  }
  if (message.startsWith("ERROR:")) {
    statusEl.textContent = message.substring(6).trim();
    statusEl.classList.add("error");
    card.classList.add("running");
    setTimeout(() => {
      card.classList.remove("running");
      if (abortBtn) abortBtn.style.display = "none";
    }, 5000);
  } else if (message === "DONE") {
    statusEl.textContent = "Done!";
    setTimeout(() => {
      card.classList.remove("running");
      if (abortBtn) abortBtn.style.display = "none";
    }, 2000);
  } else {
    statusEl.textContent = message;
  }
};




function setWireSizerMessage(message) {
  const emptyState = document.getElementById("wireSizerEmptyState");
  if (!emptyState) return;
  if (message) {
    emptyState.textContent = message;
    emptyState.hidden = false;
  } else {
    emptyState.textContent = "";
    emptyState.hidden = true;
  }
}

async function openWireSizer() {
  const dlg = document.getElementById("wireSizerDlg");
  const frame = document.getElementById("wireSizerFrame");
  if (!dlg || !frame) return;

  setWireSizerMessage("Loading Wire Sizer...");
  frame.src = "about:blank";

  if (!dlg.open) dlg.showModal();

  if (window.pywebview?.api?.get_wire_sizer_url) {
    try {
      const res = await window.pywebview.api.get_wire_sizer_url();
      if (res.status === "success" && res.url) {
        const loadTimeout = setTimeout(() => {
          setWireSizerMessage(
            "Wire Sizer is taking longer to load. Please try again."
          );
        }, 4000);
        frame.onload = () => {
          clearTimeout(loadTimeout);
          if (frame.src && !frame.src.startsWith("about:")) {
            setWireSizerMessage("");
          }
        };
        frame.onerror = () => {
          clearTimeout(loadTimeout);
          setWireSizerMessage(
            "Wire Sizer failed to load. Rebuild the tool and try again."
          );
        };
        frame.src = res.url;
        return;
      }
      setWireSizerMessage(res.message || "Wire Sizer is unavailable.");
      return;
    } catch (e) {
      setWireSizerMessage("Wire Sizer is unavailable.");
      return;
    }
  }

  setWireSizerMessage("Wire Sizer is unavailable in this environment.");
}

function closeWireSizer() {
  const dlg = document.getElementById("wireSizerDlg");
  const frame = document.getElementById("wireSizerFrame");
  if (dlg?.open) dlg.close();
  if (frame) frame.src = "about:blank";
}

// --- Circuit Breaker AI (Panel Schedule) ---

const circuitBreakerState = {
  breakerPath: "",
  directoryPath: "",
  breakerFile: null,
  directoryFile: null,
  outputMode: "new",
  newOutputPath: "",
  existingOutputPath: "",
  panelName: "",
  running: false,
};

function getCircuitBreakerOutputPath() {
  return circuitBreakerState.outputMode === "existing"
    ? circuitBreakerState.existingOutputPath
    : circuitBreakerState.newOutputPath;
}

function getCircuitBreakerFileLabel(path, file) {
  if (file?.name) return file.name;
  if (!path) return "No file selected";
  const parts = String(path).split(/[\\/]/);
  return parts[parts.length - 1] || path;
}

function setCircuitBreakerStatus(message, isError = false) {
  const statusEl = document.getElementById("cbRunStatus");
  if (!statusEl) return;
  statusEl.textContent = message || "";
  statusEl.classList.toggle("text-danger", Boolean(isError));
}

function updateCircuitBreakerUi() {
  const breakerFile = document.getElementById("cbBreakerFile");
  const directoryFile = document.getElementById("cbDirectoryFile");
  const newRow = document.getElementById("cbNewScheduleRow");
  const existingRow = document.getElementById("cbExistingScheduleRow");
  const newPath = document.getElementById("cbNewSchedulePath");
  const existingPath = document.getElementById("cbExistingSchedulePath");
  const panelNameInput = document.getElementById("cbPanelName");
  const runBtn = document.getElementById("cbRunPanelScheduleBtn");
  const modeNew = document.getElementById("cbOutputModeNew");
  const modeExisting = document.getElementById("cbOutputModeExisting");
  const breakerDrop = document.getElementById("cbBreakerDrop");
  const directoryDrop = document.getElementById("cbDirectoryDrop");
  const runningOverlay = document.getElementById("cbRunningOverlay");

  if (breakerFile) {
    const label = getCircuitBreakerFileLabel(
      circuitBreakerState.breakerPath,
      circuitBreakerState.breakerFile
    );
    breakerFile.textContent = label;
    breakerFile.dataset.empty =
      circuitBreakerState.breakerPath || circuitBreakerState.breakerFile
        ? "false"
        : "true";
  }
  if (directoryFile) {
    const label = getCircuitBreakerFileLabel(
      circuitBreakerState.directoryPath,
      circuitBreakerState.directoryFile
    );
    directoryFile.textContent = label;
    directoryFile.dataset.empty =
      circuitBreakerState.directoryPath || circuitBreakerState.directoryFile
        ? "false"
        : "true";
  }
  if (newPath) {
    const label = circuitBreakerState.newOutputPath || "No file selected";
    newPath.textContent = label;
    newPath.dataset.empty = circuitBreakerState.newOutputPath ? "false" : "true";
  }
  if (existingPath) {
    const label = circuitBreakerState.existingOutputPath || "No file selected";
    existingPath.textContent = label;
    existingPath.dataset.empty = circuitBreakerState.existingOutputPath
      ? "false"
      : "true";
  }

  if (modeNew) {
    const isActive = circuitBreakerState.outputMode === "new";
    modeNew.classList.toggle("is-active", isActive);
    modeNew.setAttribute("aria-pressed", isActive ? "true" : "false");
  }
  if (modeExisting) {
    const isActive = circuitBreakerState.outputMode === "existing";
    modeExisting.classList.toggle("is-active", isActive);
    modeExisting.setAttribute("aria-pressed", isActive ? "true" : "false");
  }

  if (newRow) newRow.hidden = circuitBreakerState.outputMode !== "new";
  if (existingRow) existingRow.hidden = circuitBreakerState.outputMode !== "existing";

  if (panelNameInput && panelNameInput.value !== circuitBreakerState.panelName) {
    panelNameInput.value = circuitBreakerState.panelName;
  }

  const ready =
    (circuitBreakerState.breakerPath || circuitBreakerState.breakerFile) &&
    (circuitBreakerState.directoryPath || circuitBreakerState.directoryFile) &&
    getCircuitBreakerOutputPath();

  if (runBtn) {
    runBtn.disabled = circuitBreakerState.running || !ready;
  }

  if (breakerDrop) {
    breakerDrop.dataset.hasFile =
      circuitBreakerState.breakerPath || circuitBreakerState.breakerFile
        ? "true"
        : "false";
    breakerDrop.classList.toggle("is-disabled", circuitBreakerState.running);
  }
  if (directoryDrop) {
    directoryDrop.dataset.hasFile =
      circuitBreakerState.directoryPath || circuitBreakerState.directoryFile
        ? "true"
        : "false";
    directoryDrop.classList.toggle("is-disabled", circuitBreakerState.running);
  }

  if (modeNew) modeNew.disabled = circuitBreakerState.running;
  if (modeExisting) modeExisting.disabled = circuitBreakerState.running;
  if (panelNameInput) panelNameInput.disabled = circuitBreakerState.running;

  if (circuitBreakerState.running && !ready) {
    circuitBreakerState.running = false;
  }
  if (runningOverlay) {
    runningOverlay.hidden = !circuitBreakerState.running;
  }

  if (circuitBreakerState.running) {
    setCircuitBreakerStatus("Running in the background...");
  } else if (ready) {
    setCircuitBreakerStatus("Ready when you are.");
  } else {
    setCircuitBreakerStatus("Select the required photos and schedule file.");
  }
}

function resetCircuitBreakerForm() {
  circuitBreakerState.breakerPath = "";
  circuitBreakerState.directoryPath = "";
  circuitBreakerState.breakerFile = null;
  circuitBreakerState.directoryFile = null;
  circuitBreakerState.outputMode = "new";
  circuitBreakerState.newOutputPath = "";
  circuitBreakerState.existingOutputPath = "";
  circuitBreakerState.panelName = "";
  circuitBreakerState.running = false;
  updateCircuitBreakerUi();
}

function setCircuitBreakerFile(kind, file) {
  if (kind === "breaker") {
    circuitBreakerState.breakerFile = file || null;
    if (file) circuitBreakerState.breakerPath = "";
  } else {
    circuitBreakerState.directoryFile = file || null;
    if (file) circuitBreakerState.directoryPath = "";
  }
  updateCircuitBreakerUi();
}

async function selectCircuitBreakerImage(kind) {
  if (!window.pywebview?.api?.select_files) {
    toast("File picker is unavailable.");
    return;
  }
  try {
    const result = await window.pywebview.api.select_files({
      allow_multiple: false,
      file_types: ["Image Files (*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tif;*.tiff)"],
    });
    if (result.status === "success" && result.paths?.length) {
      if (kind === "breaker") {
        circuitBreakerState.breakerPath = result.paths[0];
        circuitBreakerState.breakerFile = null;
      } else {
        circuitBreakerState.directoryPath = result.paths[0];
        circuitBreakerState.directoryFile = null;
      }
      updateCircuitBreakerUi();
    }
  } catch (e) {
    toast("Error selecting photo.");
  }
}

async function selectCircuitBreakerSchedulePath(mode) {
  if (!window.pywebview?.api) {
    toast("File picker is unavailable.");
    return;
  }
  if (mode === "new") {
    try {
      const selection = await window.pywebview.api.select_template_save_location(
        null,
        "Panel_Schedule",
        "xlsx"
      );
      if (selection?.status === "success" && selection.path) {
        circuitBreakerState.newOutputPath = selection.path;
        updateCircuitBreakerUi();
      }
    } catch (e) {
      toast("Error selecting save location.");
    }
    return;
  }

  try {
    const selection = await window.pywebview.api.select_files({
      allow_multiple: false,
      file_types: ["Excel Files (*.xlsx)"],
    });
    if (selection?.status === "success" && selection.paths?.length) {
      circuitBreakerState.existingOutputPath = selection.paths[0];
      updateCircuitBreakerUi();
    }
  } catch (e) {
    toast("Error selecting panel schedule.");
  }
}

async function openCircuitBreaker() {
  const dlg = document.getElementById("circuitBreakerDlg");
  if (!dlg) return;
  if (!circuitBreakerState.running) {
    resetCircuitBreakerForm();
  } else {
    circuitBreakerState.running = false;
    updateCircuitBreakerUi();
  }
  if (!dlg.open) dlg.showModal();
  updateCircuitBreakerUi();
}

function closeCircuitBreaker() {
  const dlg = document.getElementById("circuitBreakerDlg");
  if (dlg && dlg.open) dlg.close();
}

function openCircuitBreakerFilePicker(kind) {
  if (circuitBreakerState.running) return;
  const input = document.createElement("input");
  input.type = "file";
  input.accept = "image/*";
  input.onchange = () => {
    const file = input.files && input.files[0];
    if (file) setCircuitBreakerFile(kind, file);
  };
  input.click();
}

function handleCircuitBreakerDrop(kind, e, zone) {
  e.preventDefault();
  if (zone) zone.classList.remove("is-dragover");
  if (circuitBreakerState.running) return;
  const file = e.dataTransfer?.files?.[0];
  if (file) setCircuitBreakerFile(kind, file);
}

function handleCircuitBreakerDragOver(e, zone) {
  e.preventDefault();
  if (circuitBreakerState.running) return;
  if (zone) zone.classList.add("is-dragover");
}

function handleCircuitBreakerDragLeave(e, zone) {
  e.preventDefault();
  if (zone) zone.classList.remove("is-dragover");
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error || new Error("Failed to read file."));
    reader.readAsDataURL(file);
  });
}

async function fileToUploadPayload(file) {
  if (!file) return null;
  const dataUrl = await readFileAsDataUrl(file);
  return { name: file.name || "upload", dataUrl };
}

async function runCircuitBreakerInBackground() {
  const outputPath = getCircuitBreakerOutputPath();
  if (
    (!circuitBreakerState.breakerPath && !circuitBreakerState.breakerFile) ||
    (!circuitBreakerState.directoryPath && !circuitBreakerState.directoryFile)
  ) {
    toast("Select one breaker photo and one directory photo.");
    return;
  }
  if (!outputPath) {
    toast("Choose where to save or select an existing panel schedule.");
    return;
  }
  if (!window.pywebview?.api?.run_panel_schedule_background) {
    toast("Panel Schedule AI is unavailable in this environment.");
    return;
  }

  circuitBreakerState.running = true;
  updateCircuitBreakerUi();

  const payload = {
    breakerPath: circuitBreakerState.breakerPath,
    directoryPath: circuitBreakerState.directoryPath,
    breakerUploads: circuitBreakerState.breakerFile
      ? [await fileToUploadPayload(circuitBreakerState.breakerFile)]
      : [],
    directoryUploads: circuitBreakerState.directoryFile
      ? [await fileToUploadPayload(circuitBreakerState.directoryFile)]
      : [],
    outputMode: circuitBreakerState.outputMode,
    outputPath,
    panelName: circuitBreakerState.panelName?.trim() || "",
  };

  try {
    const res = await window.pywebview.api.run_panel_schedule_background(payload);
    if (res?.status === "error") {
      circuitBreakerState.running = false;
      updateCircuitBreakerUi();
      toast(res.message || "Failed to start Panel Schedule AI.");
      return;
    }
    closeCircuitBreaker();
    toast("Panel Schedule AI is running in the background.");
  } catch (e) {
    circuitBreakerState.running = false;
    updateCircuitBreakerUi();
    toast("Failed to start Panel Schedule AI.");
  }
}

window.handlePanelScheduleResult = async function (payload) {
  circuitBreakerState.running = false;
  updateCircuitBreakerUi();
  if (!payload) return;
  if (payload.status === "success") {
    toast(payload.message || "Panel Schedule AI complete.");
    const folder =
      payload.outputFolder ||
      (payload.outputPath ? payload.outputPath.split(/[\\/]/).slice(0, -1).join("\\") : "");
    if (folder && window.pywebview?.api?.open_path) {
      try {
        await window.pywebview.api.open_path(folder);
      } catch (e) {
        // Ignore open errors.
      }
    }
  } else {
    toast(payload.message || "Panel Schedule AI failed.", 6000);
  }
};

const debouncedSaveLightingSchedule = debounce(() => save(), 400);

function autoResizeTextarea(textarea) {
  if (!textarea) return;
  textarea.style.height = "auto";
  textarea.style.height = `${textarea.scrollHeight}px`;
}

function createLightingScheduleRow(seed = {}) {
  return {
    mark: seed.mark ?? "",
    description: seed.description ?? "",
    manufacturer: seed.manufacturer ?? "",
    modelNumber: seed.modelNumber ?? "",
    mounting: seed.mounting ?? "",
    volts: seed.volts ?? "",
    watts: seed.watts ?? "",
    notes: seed.notes ?? "",
  };
}

function createDefaultLightingSchedule() {
  return {
    rows: [createLightingScheduleRow()],
    generalNotes: LIGHTING_SCHEDULE_DEFAULT_GENERAL_NOTES,
    notes: LIGHTING_SCHEDULE_DEFAULT_NOTES,
  };
}

function normalizeLightingSchedule(schedule) {
  const normalized =
    schedule && typeof schedule === "object" ? schedule : {};
  if (!Array.isArray(normalized.rows) || normalized.rows.length === 0) {
    normalized.rows = [createLightingScheduleRow()];
  } else {
    normalized.rows = normalized.rows.map((row) => {
      if (!row || typeof row !== "object") {
        return createLightingScheduleRow();
      }
      LIGHTING_SCHEDULE_FIELDS.forEach((field) => {
        if (row[field] == null) row[field] = "";
      });
      return row;
    });
  }
  if (normalized.generalNotes == null) {
    normalized.generalNotes = LIGHTING_SCHEDULE_DEFAULT_GENERAL_NOTES;
  }
  if (normalized.notes == null) {
    normalized.notes = LIGHTING_SCHEDULE_DEFAULT_NOTES;
  }
  return normalized;
}

function ensureLightingSchedule(project) {
  if (!project) return { schedule: null, created: false };
  let schedule = project.lightingSchedule;
  let created = false;
  if (!schedule) {
    schedule = createDefaultLightingSchedule();
    created = true;
  }
  schedule = normalizeLightingSchedule(schedule);
  project.lightingSchedule = schedule;
  return { schedule, created };
}

function getLightingScheduleProjectLabel(project, index) {
  if (!project) return `Project ${index + 1}`;
  const id = (project.id || "").trim();
  const name = (project.name || "").trim();
  if (id && name) return `${id} - ${name}`;
  return id || name || `Project ${index + 1}`;
}

function renderLightingScheduleProjectOptions(filterText = "") {
  const select = document.getElementById("lightingScheduleProjectSelect");
  if (!select) return;
  select.innerHTML = "";
  const query = String(filterText || "").trim().toLowerCase();
  const matches = db
    .map((project, index) => ({ project, index }))
    .filter(({ project }) => {
      if (!query) return true;
      const id = String(project?.id || "").toLowerCase();
      const name = String(project?.name || "").toLowerCase();
      const nick = String(project?.nick || "").toLowerCase();
      return id.includes(query) || name.includes(query) || nick.includes(query);
    });

  matches.forEach(({ project, index }) => {
    const option = el("option", {
      value: String(index),
      textContent: getLightingScheduleProjectLabel(project, index),
    });
    select.appendChild(option);
  });
  select.disabled = matches.length === 0;

  if (!matches.length) {
    lightingScheduleProjectIndex = null;
    renderLightingSchedule(
      null,
      db.length
        ? "No projects match that search."
        : "Create a project to start a lighting fixture schedule."
    );
    return;
  }

  const availableIndexes = matches.map(({ index }) => index);
  if (availableIndexes.includes(lightingScheduleProjectIndex)) {
    select.value = String(lightingScheduleProjectIndex);
    return;
  }

  const nextIndex = availableIndexes[0];
  setLightingScheduleProject(nextIndex);
}

function getActiveLightingSchedule() {
  if (lightingScheduleProjectIndex == null) return null;
  const project = db[lightingScheduleProjectIndex];
  if (!project) return null;
  return ensureLightingSchedule(project).schedule;
}

function getLightingTemplates() {
  return Array.isArray(userSettings.lightingTemplates)
    ? userSettings.lightingTemplates
    : [];
}

function setLightingTemplates(templates) {
  userSettings.lightingTemplates = templates;
  debouncedSaveUserSettings();
}

function saveLightingTemplateFromRow(row) {
  const mark = String(row?.mark || "").trim();
  if (!mark) {
    toast("Enter a Mark before saving a template.");
    return;
  }
  const template = createLightingScheduleRow({
    ...row,
    mark,
  });
  const templates = [...getLightingTemplates()];
  const key = mark.toLowerCase();
  const existingIndex = templates.findIndex(
    (item) => String(item?.mark || "").trim().toLowerCase() === key
  );
  if (existingIndex >= 0) templates[existingIndex] = template;
  else templates.push(template);
  setLightingTemplates(templates);
  renderLightingTemplateList(lightingTemplateQuery);
  toast(`Saved template for mark ${mark}.`);
}

function addLightingTemplateToSchedule(template) {
  const schedule = getActiveLightingSchedule();
  if (!schedule) return;
  schedule.rows.push(createLightingScheduleRow(template));
  renderLightingScheduleRows(schedule);
  debouncedSaveLightingSchedule();
}

function renderLightingTemplateList(filterText = "") {
  const list = document.getElementById("lightingTemplateList");
  if (!list) return;
  list.innerHTML = "";

  const templates = getLightingTemplates();
  const query = String(filterText || "").trim().toLowerCase();
  const filtered = templates.filter((template) => {
    if (!query) return true;
    const haystack = [
      template?.mark,
      template?.description,
      template?.manufacturer,
      template?.modelNumber,
      template?.mounting,
    ]
      .map((value) => String(value || "").toLowerCase())
      .join(" ");
    return haystack.includes(query);
  });

  if (!filtered.length) {
    list.appendChild(
      el("div", {
        className: "tiny muted",
        textContent: templates.length
          ? "No templates match that search."
          : "No templates saved yet.",
      })
    );
    return;
  }

  filtered.forEach((template) => {
    const item = el("div", { className: "lighting-template-item" });
    const meta = el("div", { className: "lighting-template-meta" }, [
      el("div", {
        className: "lighting-template-mark",
        textContent: template.mark || "Untitled",
      }),
      el("div", {
        className: "lighting-template-desc",
        textContent: template.description || template.manufacturer || "",
      }),
    ]);
    const addBtn = el("button", {
      className: "btn tiny",
      textContent: "Add",
      onclick: () => addLightingTemplateToSchedule(template),
    });
    item.append(meta, addBtn);
    list.appendChild(item);
  });
}

function renderLightingScheduleRows(schedule) {
  const body = document.getElementById("lightingScheduleRows");
  if (!body) return;
  body.innerHTML = "";
  schedule.rows.forEach((row, rowIndex) => {
    const tr = document.createElement("tr");
    LIGHTING_SCHEDULE_FIELDS.forEach((field) => {
      const td = document.createElement("td");
      const input = document.createElement("textarea");
      input.rows = 1;
      input.value = row[field] ?? "";
      input.addEventListener("input", (e) => {
        row[field] = e.target.value;
        autoResizeTextarea(input);
        debouncedSaveLightingSchedule();
      });
      requestAnimationFrame(() => autoResizeTextarea(input));
      if (field === "notes") {
        const wrap = document.createElement("div");
        wrap.className = "lighting-schedule-row-notes";
        wrap.appendChild(input);
        const actions = document.createElement("div");
        actions.className = "lighting-schedule-row-actions";
        const saveBtn = document.createElement("button");
        saveBtn.type = "button";
        saveBtn.className = "lighting-schedule-template-save";
        saveBtn.textContent = "S";
        saveBtn.title = "Save template";
        saveBtn.setAttribute("aria-label", "Save template");
        saveBtn.addEventListener("click", () =>
          saveLightingTemplateFromRow(row)
        );
        const removeBtn = document.createElement("button");
        removeBtn.type = "button";
        removeBtn.className = "lighting-schedule-remove";
        removeBtn.textContent = "X";
        removeBtn.title = "Remove row";
        removeBtn.setAttribute("aria-label", "Remove row");
        removeBtn.addEventListener("click", () =>
          removeLightingScheduleRow(rowIndex)
        );
        actions.append(saveBtn, removeBtn);
        wrap.appendChild(actions);
        td.appendChild(wrap);
      } else {
        td.appendChild(input);
      }
      tr.appendChild(td);
    });
    body.appendChild(tr);
  });
}

function renderLightingSchedule(
  schedule,
  emptyMessage = "Create a project to start a lighting fixture schedule."
) {
  const emptyState = document.getElementById("lightingScheduleEmptyState");
  const tableWrap = document.getElementById("lightingScheduleTableWrap");
  const addRowBtn = document.getElementById("lightingScheduleAddRow");
  const generalNotes = document.getElementById("lightingScheduleGeneralNotes");
  const notes = document.getElementById("lightingScheduleNotes");

  if (!schedule) {
    if (emptyState) {
      emptyState.textContent = emptyMessage;
      emptyState.hidden = false;
    }
    if (tableWrap) tableWrap.hidden = true;
    if (addRowBtn) addRowBtn.disabled = true;
    if (generalNotes) {
      generalNotes.value = "";
      autoResizeTextarea(generalNotes);
    }
    if (notes) {
      notes.value = "";
      autoResizeTextarea(notes);
    }
    return;
  }

  if (emptyState) emptyState.hidden = true;
  if (tableWrap) tableWrap.hidden = false;
  if (addRowBtn) addRowBtn.disabled = false;
  if (generalNotes) {
    generalNotes.value = schedule.generalNotes ?? "";
    autoResizeTextarea(generalNotes);
  }
  if (notes) {
    notes.value = schedule.notes ?? "";
    autoResizeTextarea(notes);
  }
  renderLightingScheduleRows(schedule);
}

function setLightingScheduleProject(index) {
  if (!Number.isInteger(index) || !db[index]) {
    lightingScheduleProjectIndex = null;
    renderLightingSchedule(null);
    return;
  }
  lightingScheduleProjectIndex = index;
  const select = document.getElementById("lightingScheduleProjectSelect");
  if (select) select.value = String(index);
  const { schedule, created } = ensureLightingSchedule(db[index]);
  if (created) save();
  renderLightingSchedule(schedule);
}

function addLightingScheduleRow() {
  const schedule = getActiveLightingSchedule();
  if (!schedule) return;
  schedule.rows.push(createLightingScheduleRow());
  renderLightingScheduleRows(schedule);
  debouncedSaveLightingSchedule();
}

function removeLightingScheduleRow(rowIndex) {
  const schedule = getActiveLightingSchedule();
  if (!schedule) return;
  if (schedule.rows.length <= 1) {
    schedule.rows[0] = createLightingScheduleRow();
  } else {
    schedule.rows.splice(rowIndex, 1);
  }
  renderLightingScheduleRows(schedule);
  debouncedSaveLightingSchedule();
}

function openLightingSchedule() {
  const dlg = document.getElementById("lightingScheduleDlg");
  if (!dlg) return;
  const projectSearch = document.getElementById(
    "lightingScheduleProjectSearch"
  );
  if (projectSearch) projectSearch.value = lightingScheduleProjectQuery;
  renderLightingScheduleProjectOptions(lightingScheduleProjectQuery);

  const templateSearch = document.getElementById("lightingTemplateSearch");
  if (templateSearch) templateSearch.value = lightingTemplateQuery;
  renderLightingTemplateList(lightingTemplateQuery);
  if (!dlg.open) dlg.showModal();
}

function closeLightingSchedule() {
  const dlg = document.getElementById("lightingScheduleDlg");
  if (dlg?.open) dlg.close();
  save();
}

// ===================== ONBOARDING SYSTEM =====================

let currentOnboardingStep = 1;
const totalOnboardingSteps = 4;

function isNewUser() {
  return !userSettings.userName || userSettings.userName.trim() === "";
}

function showOnboardingModal() {
  currentOnboardingStep = 1;
  updateOnboardingUI();
  document.getElementById("onboardingDlg").showModal();
}

function skipOnboarding() {
  closeDlg("onboardingDlg");
  showMainApp();
}

function updateOnboardingUI() {
  // Update progress indicators
  document.querySelectorAll(".progress-step").forEach((step, index) => {
    step.classList.toggle("active", index + 1 <= currentOnboardingStep);
  });

  // Update step visibility
  document.querySelectorAll(".onboarding-step").forEach((step, index) => {
    step.classList.toggle("active", index + 1 === currentOnboardingStep);
  });

  // Update title
  const titles = [
    "Welcome to ACIES!",
    "What's your full name?",
    "What's your engineering discipline?",
    "Set up your AI assistant",
  ];
  document.getElementById("onboardingTitle").textContent =
    titles[currentOnboardingStep - 1];

  // Update navigation buttons
  const prevBtn = document.getElementById("onboardingPrevBtn");
  const nextBtn = document.getElementById("onboardingNextBtn");

  prevBtn.disabled = currentOnboardingStep === 1;

  if (currentOnboardingStep === totalOnboardingSteps) {
    nextBtn.textContent = "Get Started";
  } else {
    nextBtn.textContent = "Next";
  }

  // Add Enter key navigation for quick advancement
  setupOnboardingEnterKeyListeners();
}

function setupOnboardingEnterKeyListeners() {
  // Remove existing listeners to avoid duplicates
  const modal = document.getElementById("onboardingDlg");
  const nameInput = document.getElementById("onboarding_userName");
  const apiKeyInput = document.getElementById("onboarding_apiKey");

  // Remove previous listeners
  modal.removeEventListener("keydown", handleOnboardingModalKeydown);
  if (nameInput)
    nameInput.removeEventListener("keydown", handleOnboardingInputKeydown);
  if (apiKeyInput)
    apiKeyInput.removeEventListener("keydown", handleOnboardingInputKeydown);

  // Add listeners based on current step
  if (currentOnboardingStep === 1) {
    // Step 1: Welcome - Enter on modal to proceed
    modal.addEventListener("keydown", handleOnboardingModalKeydown);
  } else if (currentOnboardingStep === 2) {
    // Step 2: Name input - Enter to validate and proceed
    if (nameInput)
      nameInput.addEventListener("keydown", handleOnboardingInputKeydown);
  } else if (currentOnboardingStep === 3) {
    // Step 3: Discipline checkboxes - Enter on modal to validate and proceed
    modal.addEventListener("keydown", handleOnboardingModalKeydown);
  } else if (currentOnboardingStep === 4) {
    // Step 4: API key input - Enter to validate and proceed
    if (apiKeyInput)
      apiKeyInput.addEventListener("keydown", handleOnboardingInputKeydown);
  }
}

function handleOnboardingModalKeydown(e) {
  if (e.key === "Enter") {
    e.preventDefault();
    nextOnboardingStep();
  }
}

function handleOnboardingInputKeydown(e) {
  if (e.key === "Enter") {
    e.preventDefault();
    nextOnboardingStep();
  }
}

function nextOnboardingStep() {
  if (!validateCurrentStep()) return;

  if (currentOnboardingStep < totalOnboardingSteps) {
    currentOnboardingStep++;
    updateOnboardingUI();
  } else {
    completeOnboarding();
  }
}

function prevOnboardingStep() {
  if (currentOnboardingStep > 1) {
    currentOnboardingStep--;
    updateOnboardingUI();
  }
}

function validateCurrentStep() {
  switch (currentOnboardingStep) {
    case 2: // Name step
      const name = val("onboarding_userName").trim();
      if (!name) {
        toast("Please enter your full name.");
        return false;
      }
      userSettings.userName = name;
      break;
    case 3: // Discipline step
      const selectedDisciplines = document.querySelectorAll(
        'input[name="onboarding_discipline_checkbox"]:checked'
      );
      if (selectedDisciplines.length === 0) {
        toast("Please select at least one engineering discipline.");
        return false;
      }
      userSettings.discipline = Array.from(selectedDisciplines).map(
        (cb) => cb.value
      );
      break;
    case 4: // API Key step
      const apiKey = val("onboarding_apiKey").trim();
      if (!apiKey) {
        toast("Please enter your Google Gemini API key.");
        return false;
      }
      userSettings.apiKey = apiKey;
      break;
  }
  return true;
}

async function completeOnboarding() {
  try {
    await window.pywebview.api.save_user_settings(userSettings);
    closeDlg("onboardingDlg");
    showMainApp();
    toast("Welcome to ACIES! Your settings have been saved.");
  } catch (e) {
    toast("⚠️ Could not save settings. Please try again.");
  }
}

function showMainApp() {
  document.querySelector("main").style.display = "block";
  document.querySelector(".app-header").style.display = "flex";
}

function hideMainApp() {
  document.querySelector("main").style.display = "none";
  document.querySelector(".app-header").style.display = "none";
}

function showAppLoader() {
  const loader = document.getElementById("appLoader");
  if (!loader) return;
  loader.classList.remove("hidden");
  loader.removeAttribute("aria-hidden");
  document.body.classList.add("app-loading");
}

function hideAppLoader() {
  const loader = document.getElementById("appLoader");
  if (!loader) return;
  loader.classList.add("hidden");
  loader.setAttribute("aria-hidden", "true");
  document.body.classList.remove("app-loading");
}

function updateAggregationOptions() {
  const allowed = allowedAggregations[currentStatsTimespan] || [];
  document.querySelectorAll("#statsAggGroup .filter-chip").forEach((btn) => {
    const agg = btn.dataset.agg;
    if (allowed.includes(agg)) {
      btn.removeAttribute("disabled");
      btn.classList.remove("disabled");
    } else {
      btn.setAttribute("disabled", "true");
      btn.classList.add("disabled");
    }
  });
  // If current aggregation not allowed, switch to first allowed
  if (!allowed.includes(currentStatsAggregation)) {
    currentStatsAggregation = allowed[0];
    // Update active class
    document
      .querySelectorAll("#statsAggGroup .filter-chip")
      .forEach((b) => b.classList.remove("active"));
    document
      .querySelector(
        `#statsAggGroup .filter-chip[data-agg="${currentStatsAggregation}"]`
      )
      ?.classList.add("active");
  }
}

function showStatsModal() {
  document.getElementById("statsDlg").showModal();
  updateAggregationOptions();
  renderStats();
}

function getTimespanDateRange(timespan) {
  const now = new Date();
  const start = new Date(now);

  switch (timespan) {
    case "1W":
      start.setDate(now.getDate() - 7);
      break;
    case "1M":
      start.setMonth(now.getMonth() - 1);
      break;
    case "1Y":
      start.setFullYear(now.getFullYear() - 1);
      break;
    case "3Y":
      start.setFullYear(now.getFullYear() - 3);
      break;
    case "all":
      // Find the earliest date in the DB
      {
        let earliest = now;
        getAllDeliverables().forEach(({ deliverable }) => {
          const d = parseDueStr(deliverable?.due);
          if (d && d < earliest) earliest = d;
        });
        return { start: earliest, end: now };
      }
  }

  return { start, end: now };
}

function renderStats() {
  const { start, end } = getTimespanDateRange(currentStatsTimespan);
  const allDeliverables = getAllDeliverables().map(
    ({ deliverable }) => deliverable
  );

  // Filter deliverables within the timespan
  const filteredDeliverables = allDeliverables.filter((d) => {
    const dueDate = parseDueStr(d.due);
    return dueDate && dueDate >= start && dueDate <= end;
  });

  // Calculate stats from entire history
  const total = allDeliverables.length;
  const completed = allDeliverables.filter((d) => isFinished(d)).length;

  // Calculate averages for weeks/months with completions
  const allCompletedDeliverables = allDeliverables
    .filter((d) => isFinished(d))
    .map((d) => ({
      date: parseDueStr(d.due),
      ...d,
    }))
    .filter((d) => d.date);

  // Group completions by week
  const weeklyCompletions = {};
  allCompletedDeliverables.forEach((d) => {
    const weekStart = new Date(d.date);
    weekStart.setDate(d.date.getDate() - d.date.getDay());
    weekStart.setHours(0, 0, 0, 0);
    const weekKey = weekStart.getTime();
    weeklyCompletions[weekKey] = (weeklyCompletions[weekKey] || 0) + 1;
  });

  // Group completions by month
  const monthlyCompletions = {};
  allCompletedDeliverables.forEach((d) => {
    const monthKey = `${d.date.getFullYear()}-${String(
      d.date.getMonth() + 1
    ).padStart(2, "0")}`;
    monthlyCompletions[monthKey] = (monthlyCompletions[monthKey] || 0) + 1;
  });

  // Calculate averages for active periods
  const activeWeeks = Object.values(weeklyCompletions);
  const avgPerWeek =
    activeWeeks.length > 0
      ? (
        activeWeeks.reduce((sum, count) => sum + count, 0) /
        activeWeeks.length
      ).toFixed(1)
      : "0";

  const activeMonths = Object.values(monthlyCompletions);
  const avgPerMonth =
    activeMonths.length > 0
      ? (
        activeMonths.reduce((sum, count) => sum + count, 0) /
        activeMonths.length
      ).toFixed(1)
      : "0";

  // Calculate due stats
  const now = new Date();
  const weekStart = new Date(now);
  weekStart.setDate(now.getDate() - now.getDay());
  const weekEnd = new Date(weekStart);
  weekEnd.setDate(weekStart.getDate() + 6);

  let dueThisWeek = 0,
    dueLastWeek = 0,
    upcoming = 0;

  // Use global DB for forecast stats, not filtered
  allDeliverables.forEach((d) => {
    const due = parseDueStr(d.due);
    if (due) {
      if (due >= weekStart && due <= weekEnd) dueThisWeek++;
      else if (
        due >= new Date(weekStart.getTime() - 7 * 24 * 60 * 60 * 1000) &&
        due < weekStart
      )
        dueLastWeek++;
      else if (due > weekEnd) upcoming++;
    }
  });

  // Calculate year stats (global)
  const currentYear = now.getFullYear();
  const lastYear = currentYear - 1;
  let completedThisYear = 0,
    completedLastYear = 0;

  allDeliverables.forEach((d) => {
    if (isFinished(d)) {
      const due = parseDueStr(d.due);
      if (due) {
        if (due.getFullYear() === currentYear) completedThisYear++;
        if (due.getFullYear() === lastYear) completedLastYear++;
      }
    }
  });

  // Update numbers view
  setStat("statTotal", total);
  setStat("statCompleted", completed);
  setStat("statAvg", avgPerWeek);
  setStat("statAvgMonth", avgPerMonth);
  setStat("statLastWeek", dueLastWeek);
  setStat("statThisWeek", dueThisWeek);
  setStat("statUpcoming", upcoming);
  setStat("statCompletedLastYear", completedLastYear);
  setStat("statCompletedThisYear", completedThisYear);

  // Always render chart with explicit aggregation
  renderStatsChart(filteredDeliverables, start, end, currentStatsAggregation);
}

function renderStatsChart(projects, start, end, aggregation) {
  const canvas = document.getElementById("statsChart");
  const ctx = canvas.getContext("2d");

  // Destroy existing chart if any
  if (window.statsChartInstance) {
    window.statsChartInstance.destroy();
  }

  // Get completed projects with dates
  const completedProjects = projects
    .filter((p) => isFinished(p))
    .map((p) => ({ date: parseDueStr(p.due), ...p }))
    .filter((p) => p.date);

  // Helper function to get period key for a date
  function getPeriodKey(date, aggregationType) {
    if (aggregationType === "day") {
      // Use ISO date format for consistency
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, "0");
      const day = String(date.getDate()).padStart(2, "0");
      return `${year}-${month}-${day}`;
    } else if (aggregationType === "week") {
      // Get the start of the week (Sunday)
      const d = new Date(date);
      const day = d.getDay();
      const diff = d.getDate() - day;
      const weekStart = new Date(d.setDate(diff));
      weekStart.setHours(0, 0, 0, 0);
      const year = weekStart.getFullYear();
      const month = String(weekStart.getMonth() + 1).padStart(2, "0");
      const dayOfMonth = String(weekStart.getDate()).padStart(2, "0");
      return `${year}-${month}-${dayOfMonth}`;
    } else if (aggregationType === "month") {
      return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(
        2,
        "0"
      )}`;
    } else if (aggregationType === "quarter") {
      const year = date.getFullYear();
      const quarter = Math.floor(date.getMonth() / 3) + 1;
      return `${year}-Q${quarter}`;
    }
  }

  // Helper function to format period label
  function formatPeriodLabel(key, aggregationType) {
    if (aggregationType === "day") {
      const [year, month, day] = key.split("-");
      const date = new Date(year, parseInt(month) - 1, parseInt(day));
      return date.toLocaleDateString();
    } else if (aggregationType === "week") {
      const [year, month, day] = key.split("-");
      const weekStart = new Date(year, parseInt(month) - 1, parseInt(day));
      return `Week of ${weekStart.toLocaleDateString(undefined, {
        month: "short",
        day: "numeric",
      })}`;
    } else if (aggregationType === "month") {
      const [year, month] = key.split("-");
      const date = new Date(year, parseInt(month) - 1, 1);
      return date.toLocaleDateString(undefined, {
        year: "numeric",
        month: "short",
      });
    } else if (aggregationType === "quarter") {
      const [year, q] = key.split("-Q");
      return `${year} Q${q}`;
    }
  }

  // Aggregate projects by period
  const periodCounts = {};
  completedProjects.forEach((p) => {
    const periodKey = getPeriodKey(p.date, aggregation);
    periodCounts[periodKey] = (periodCounts[periodKey] || 0) + 1;
  });

  // Generate all periods in the range
  const labels = [];
  const data = [];
  const current = new Date(start);
  current.setHours(0, 0, 0, 0);

  if (aggregation === "day") {
    while (current <= end) {
      const year = current.getFullYear();
      const month = String(current.getMonth() + 1).padStart(2, "0");
      const day = String(current.getDate()).padStart(2, "0");
      const key = `${year}-${month}-${day}`;
      labels.push(formatPeriodLabel(key, aggregation));
      data.push(periodCounts[key] || 0);
      current.setDate(current.getDate() + 1);
    }
  } else if (aggregation === "week") {
    // Start from the beginning of the week
    const day = current.getDay();
    const diff = current.getDate() - day;
    current.setDate(diff);

    while (current <= end) {
      const year = current.getFullYear();
      const month = String(current.getMonth() + 1).padStart(2, "0");
      const dayOfMonth = String(current.getDate()).padStart(2, "0");
      const key = `${year}-${month}-${dayOfMonth}`;
      labels.push(formatPeriodLabel(key, aggregation));
      data.push(periodCounts[key] || 0);
      current.setDate(current.getDate() + 7);
    }
  } else if (aggregation === "month") {
    // Start from the beginning of the month
    current.setDate(1);

    while (current <= end) {
      const key = `${current.getFullYear()}-${String(
        current.getMonth() + 1
      ).padStart(2, "0")}`;
      labels.push(formatPeriodLabel(key, aggregation));
      data.push(periodCounts[key] || 0);
      current.setMonth(current.getMonth() + 1);
    }
  } else if (aggregation === "quarter") {
    // Start from the beginning of the quarter
    const currentQuarter = Math.floor(current.getMonth() / 3);
    current.setMonth(currentQuarter * 3, 1);

    while (current <= end) {
      const year = current.getFullYear();
      const quarter = Math.floor(current.getMonth() / 3) + 1;
      const key = `${year}-Q${quarter}`;
      labels.push(formatPeriodLabel(key, aggregation));
      data.push(periodCounts[key] || 0);
      current.setMonth(current.getMonth() + 3);
    }
  }

  // Get current theme colors
  const style = getComputedStyle(document.body);
  const textColor = style.getPropertyValue("--text").trim();
  const gridColor = style.getPropertyValue("--glass-border").trim();
  const surfaceColor = style.getPropertyValue("--bg-secondary").trim(); // Use bg-secondary for better opacity

  window.statsChartInstance = new Chart(ctx, {
    type: "bar",
    data: {
      labels: labels,
      datasets: [
        {
          label: "Deliverables Completed",
          data: data,
          backgroundColor: "rgba(16, 185, 129, 0.6)",
          borderColor: "var(--accent)",
          borderWidth: 1,
          borderRadius: 4,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          display: false,
        },
        tooltip: {
          backgroundColor: surfaceColor,
          titleColor: textColor,
          bodyColor: textColor,
          borderColor: gridColor,
          borderWidth: 1,
          padding: 10,
          displayColors: false,
          callbacks: {
            label: function (context) {
              const count = context.parsed.y;
              return `${count} deliverable${count !== 1 ? "s" : ""} completed`;
            },
          },
        },
      },
      scales: {
        x: {
          title: {
            display: true,
            text: "Time Period",
            color: textColor,
          },
          grid: {
            color: gridColor,
            display: false,
          },
          ticks: {
            color: textColor,
            maxTicksLimit: 12,
            maxRotation: 45,
            minRotation: 0,
          },
        },
        y: {
          title: {
            display: true,
            text: "Deliverables Completed",
            color: textColor,
          },
          beginAtZero: true,
          grid: {
            color: gridColor,
          },
          ticks: {
            color: textColor,
            stepSize: 1,
            precision: 0,
          },
        },
      },
    },
  });
}
function showApiKeyHelp() {
  document.getElementById("apiKeyHelpDlg").showModal();
}

// ===================== SETUP HELP BANNER =====================

function showSetupHelpBanner() {
  const banner = document.getElementById("setupHelpBanner");
  if (banner && userSettings.showSetupHelp !== false) {
    banner.style.display = "block";
  }
}

function hideSetupHelpBanner() {
  const banner = document.getElementById("setupHelpBanner");
  if (banner) {
    banner.style.display = "none";
  }
}

function startSetupHelp() {
  hideSetupHelpBanner();
  // Show the full onboarding experience
  currentOnboardingStep = 1;
  updateOnboardingUI();
  document.getElementById("onboardingDlg").showModal();
}

async function dismissSetupHelp(type) {
  hideSetupHelpBanner();

  if (type === "never") {
    userSettings.showSetupHelp = false;
    try {
      await window.pywebview.api.save_user_settings(userSettings);
    } catch (e) {
      console.warn("Could not save setup help preference:", e);
    }
  }
  // For "later", we don't change the setting, so it will show again next time
}

// ===================== INITIALIZATION & EVENTS =====================

function initTabbedInterfaces() {
  const mainTabContainer = document.querySelector(".main-nav");
  const notesResults = document.getElementById("notesSearchResults");

  mainTabContainer.addEventListener("click", (e) => {
    if (!e.target.matches(".main-tab-btn")) return;
    const tab = e.target.dataset.tab;

    document
      .querySelectorAll(".main-tab-btn")
      .forEach((b) => b.classList.toggle("active", b.dataset.tab === tab));
    document.querySelectorAll(".tab-panel").forEach((p) => {
      p.hidden = p.id !== `${tab}-panel`;
      p.classList.toggle("active", p.id === `${tab}-panel`);
    });

    if (notesResults) notesResults.innerHTML = "";

    if (tab === "plugins") {
      ensureBundlesRendered();
    } else if (tab === "projects") {
      render();
    } else if (tab === "timesheets") {
      renderTimesheets();
    } else if (tab === "templates") {
      renderTemplates();
    }
  });
}

function initEventListeners() {
  document.getElementById("search").addEventListener(
    "input",
    debounce(() => render(), 250)
  );
  document.getElementById("notesSearch").addEventListener(
    "input",
    debounce(() => renderNoteSearchResults(), 250)
  );

  document.getElementById("mainHelpBtn").onclick = () =>
    openExternalUrl(HELP_LINKS.main);
  document.getElementById("projectsHelpBtn").onclick = () =>
    openExternalUrl(HELP_LINKS.projects);

  document.querySelectorAll(".tool-card-settings").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      const targetId = btn.dataset.settingsTarget;
      if (!targetId) return;
      if (targetId === "publishSettingsDlg") syncPublishOptionsInputs();
      if (targetId === "freezeSettingsDlg") syncFreezeOptionsInputs();
      if (targetId === "thawSettingsDlg") syncThawOptionsInputs();
      if (targetId === "cleanSettingsDlg") syncCleanOptionsInputs();
      const dlg = document.getElementById(targetId);
      if (dlg) dlg.showModal();
    });
  });

  const handleAppFocus = () => checkBundlesForUpdates({ showIndicator: true });
  window.addEventListener("focus", handleAppFocus);
  document.addEventListener("visibilitychange", () => {
    if (!document.hidden) handleAppFocus();
  });

  const toggleNonPrimaryBtn = document.getElementById("toggleNonPrimaryBtn");
  const updateToggleNonPrimaryState = () => {
    if (!toggleNonPrimaryBtn) return;
    toggleNonPrimaryBtn.classList.toggle("is-active", hideNonPrimary);
    toggleNonPrimaryBtn.setAttribute("aria-pressed", String(hideNonPrimary));
    toggleNonPrimaryBtn.title = hideNonPrimary
      ? "Show all deliverables"
      : "Hide non-primary deliverables";
  };
  if (toggleNonPrimaryBtn) {
    updateToggleNonPrimaryState();
    toggleNonPrimaryBtn.onclick = () => {
      hideNonPrimary = !hideNonPrimary;
      updateToggleNonPrimaryState();
      render();
    };
  }
  document.getElementById("notesHelpBtn").onclick = () =>
    openExternalUrl(HELP_LINKS.notes);
  document.getElementById("toolsHelpBtn").onclick = () =>
    openExternalUrl(HELP_LINKS.tools);
  document.getElementById("pluginsHelpBtn").onclick = () =>
    openExternalUrl(HELP_LINKS.plugins);
  document.getElementById("timesheetsHelpBtn").onclick = () =>
    openExternalUrl(HELP_LINKS.timesheets);

  // Templates event listeners
  const addTemplateBtn = document.getElementById("addTemplateBtn");
  if (addTemplateBtn) addTemplateBtn.onclick = () => handleAddTemplate();
  const addTemplateEmptyBtn = document.getElementById("addTemplateEmptyBtn");
  if (addTemplateEmptyBtn) addTemplateEmptyBtn.onclick = () => handleAddTemplate();
  const browseTemplateBtn = document.getElementById("btnBrowseTemplateFile");
  if (browseTemplateBtn) browseTemplateBtn.onclick = () => handleBrowseTemplateFile();
  const saveNewTemplateBtn = document.getElementById("btnSaveNewTemplate");
  if (saveNewTemplateBtn) saveNewTemplateBtn.onclick = () => handleSaveNewTemplate();
  const browseDestinationBtn = document.getElementById("btnBrowseDestination");
  if (browseDestinationBtn) browseDestinationBtn.onclick = () => handleBrowseDestination();
  const executeCopyBtn = document.getElementById("btnExecuteCopy");
  if (executeCopyBtn) executeCopyBtn.onclick = () => handleExecuteCopy();
  const copyDeliverableSelect = document.getElementById("copy_deliverable_select");
  if (copyDeliverableSelect) {
    copyDeliverableSelect.onchange = () => updateCopyDialogAutoFields();
  }
  const templatesHelpBtn = document.getElementById("templatesHelpBtn");
  if (templatesHelpBtn) {
    templatesHelpBtn.onclick = () => openExternalUrl(HELP_LINKS.main);
  }

  document.getElementById("prevWeekBtn").onclick = () =>
    navigateTimesheetWeek(-1);
  document.getElementById("nextWeekBtn").onclick = () =>
    navigateTimesheetWeek(1);
  document.getElementById("currentWeekBtn").onclick = () =>
    goToCurrentWeek();
  document.getElementById("addTimesheetEntryBtn").onclick = () =>
    openAddTimesheetProjectDialog();
  document.getElementById("addManualTimesheetEntryBtn").onclick = () =>
    addManualTimesheetEntry();
  document.getElementById("exportTimesheetBtn").onclick = () =>
    exportTimesheetToExcel();

  // Expense sheet event listeners
  document.getElementById("addExpenseProjectBtn").onclick = () =>
    openAddExpenseProjectDialog();
  document.getElementById("btnSaveExpenseEntry").onclick = () =>
    saveExpenseEntry();

  document.getElementById("checkUpdateBtn").onclick = () =>
    refreshAppUpdateStatus({ manual: true });
  document.getElementById("appUpdateBtn").onclick = installAppUpdate;
  document.getElementById("themeToggleBtn").onclick = () => {
    const currentTheme =
      document.documentElement.getAttribute("data-theme") || "dark";
    const nextTheme = currentTheme === "light" ? "dark" : "light";
    persistThemePreference(nextTheme);
  };

  const mergeSimilarBtn = document.getElementById("btnMergeSimilarProjects");
  if (mergeSimilarBtn) mergeSimilarBtn.onclick = scanAndMergeSimilarProjects;

  document.getElementById("quickNew").onclick = openNew;
  document.getElementById("settingsBtn").onclick = async () => {
    hideSetupHelpBanner(); // Hide the banner when user manually opens settings
    await populateSettingsModal();
    document.getElementById("settingsDlg").showModal();
  };
  document.getElementById("statsBtn").onclick = () => showStatsModal();
  document.getElementById("settings_howToSetupBtn").onclick = () =>
    document.getElementById("apiKeyHelpDlg").showModal();
  const openTimesheetsBtn = document.getElementById(
    "settings_openTimesheetsFolder"
  );
  if (openTimesheetsBtn) {
    openTimesheetsBtn.onclick = async () => {
      try {
        const res = await window.pywebview.api.open_timesheets_folder();
        if (res && res.status !== "success") {
          throw new Error(res.message || "Failed to open timesheets folder.");
        }
      } catch (e) {
        console.warn("Failed to open timesheets folder:", e);
        toast("Couldn't open timesheets folder.");
      }
    };
  }
  const refreshTimesheetsBtn = document.getElementById(
    "settings_refreshTimesheetsInfo"
  );
  if (refreshTimesheetsBtn) {
    refreshTimesheetsBtn.onclick = async () => {
      await refreshTimesheetsInfo();
    };
  }
  const forceSaveTimesheetsBtn = document.getElementById(
    "settings_forceSaveTimesheets"
  );
  if (forceSaveTimesheetsBtn) {
    forceSaveTimesheetsBtn.onclick = async () => {
      try {
        await saveTimesheets();
        await refreshTimesheetsInfo();
        toast("Timesheets saved.");
      } catch (e) {
        console.warn("Failed to force save timesheets:", e);
        toast("Failed to save timesheets.");
      }
    };
  }
  const writeTimesheetsTestBtn = document.getElementById(
    "settings_writeTimesheetsTest"
  );
  if (writeTimesheetsTestBtn) {
    writeTimesheetsTestBtn.onclick = async () => {
      try {
        const res = await window.pywebview.api.write_timesheets_test_file();
        if (res && res.status !== "success") {
          throw new Error(
            res.message ||
            `Failed to write test file: ${res.path || "unknown path"}`
          );
        }
        await refreshTimesheetsInfo();
        const pathInfo = res?.path ? ` ${res.path}` : "";
        toast(`Test file written:${pathInfo}`);
      } catch (e) {
        console.warn("Failed to write timesheets test file:", e);
        toast("Failed to write test file.");
      }
    };
  }
  document.getElementById("btnStartSetupGuide").onclick = () => {
    closeDlg("settingsDlg");
    startSetupHelp();
  };

  // Statistics modal controls
  const statsRangeGroup = document.getElementById("statsRangeGroup");
  if (statsRangeGroup) {
    statsRangeGroup.addEventListener("click", (e) => {
      if (e.target.matches(".filter-chip")) {
        currentStatsTimespan = e.target.dataset.range;
        // Update active class
        document
          .querySelectorAll("#statsRangeGroup .filter-chip")
          .forEach((b) => b.classList.remove("active"));
        e.target.classList.add("active");
        updateAggregationOptions();
        renderStats();
      }
    });
  }

  const statsAggGroup = document.getElementById("statsAggGroup");
  if (statsAggGroup) {
    statsAggGroup.addEventListener("click", (e) => {
      if (
        e.target.matches(".filter-chip") &&
        !e.target.hasAttribute("disabled")
      ) {
        currentStatsAggregation = e.target.dataset.agg;
        // Update active class
        document
          .querySelectorAll("#statsAggGroup .filter-chip")
          .forEach((b) => b.classList.remove("active"));
        e.target.classList.add("active");
        renderStats();
      }
    });
  }

  document.getElementById("btnSaveProject").onclick = onSaveProject;

  const pathInput = document.getElementById("f_path");
  if (pathInput) {
    pathInput.addEventListener("input", () =>
      applyProjectFromPath(pathInput.value)
    );
    pathInput.addEventListener("paste", () => {
      setTimeout(() => applyProjectFromPath(pathInput.value, { force: true }), 0);
    });
  }

  document.getElementById("dueFilterGroup").addEventListener("click", (e) => {
    if (e.target.matches(".filter-chip")) {
      dueFilter = e.target.dataset.dueFilter;
      document
        .querySelectorAll("#dueFilterGroup .filter-chip")
        .forEach((b) => b.classList.remove("active"));
      e.target.classList.add("active");
      render();
    }
  });
  document
    .getElementById("statusFilterGroup")
    .addEventListener("click", (e) => {
      if (e.target.matches(".filter-chip")) {
        statusFilter = e.target.dataset.filter;
        document
          .querySelectorAll("#statusFilterGroup .filter-chip")
          .forEach((b) => b.classList.remove("active"));
        e.target.classList.add("active");
        render();
      }
    });

  document
    .getElementById("toolPublishDwgs")
    .addEventListener("click", async (e) => {
      if (e.currentTarget.classList.contains("running")) return;
      if (!userSettings.autocadPath) {
        await showAutocadSelectModal();
        return;
      }
      e.currentTarget.classList.add("running");
      window.updateToolStatus("toolPublishDwgs", "Initializing...");
      await window.pywebview.api.run_publish_script();
    });

  document
    .getElementById("toolFreezeLayers")
    .addEventListener("click", async (e) => {
      if (e.currentTarget.classList.contains("running")) return;
      if (!userSettings.autocadPath) {
        await showAutocadSelectModal();
        return;
      }
      e.currentTarget.classList.add("running");
      window.updateToolStatus("toolFreezeLayers", "Initializing...");
      await window.pywebview.api.run_freeze_layers_script();
    });

  document
    .getElementById("toolThawLayers")
    .addEventListener("click", async (e) => {
      if (e.currentTarget.classList.contains("running")) return;
      if (!userSettings.autocadPath) {
        await showAutocadSelectModal();
        return;
      }
      e.currentTarget.classList.add("running");
      window.updateToolStatus("toolThawLayers", "Initializing...");
      await window.pywebview.api.run_thaw_layers_script();
    });

  document
    .getElementById("toolCleanXrefs")
    .addEventListener("click", async (e) => {
      if (e.currentTarget.classList.contains("running")) return;
      if (!userSettings.autocadPath) {
        await showAutocadSelectModal();
        return;
      }
      e.currentTarget.classList.add("running");
      window.updateToolStatus("toolCleanXrefs", "Initializing...");
      await window.pywebview.api.run_clean_xrefs_script();
    });

  const narrativeTemplateBtn = document.getElementById(
    "toolCreateNarrativeTemplate"
  );
  if (narrativeTemplateBtn) {
    const handler = () =>
      handleTemplateToolSave("narrative", "Narrative of Changes");
    narrativeTemplateBtn.addEventListener("click", handler);
    narrativeTemplateBtn.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        handler();
      }
    });
  }

  const planCheckTemplateBtn = document.getElementById(
    "toolCreatePlanCheckTemplate"
  );
  if (planCheckTemplateBtn) {
    const handler = () =>
      handleTemplateToolSave("planCheck", "Plan Check Comments");
    planCheckTemplateBtn.addEventListener("click", handler);
    planCheckTemplateBtn.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        handler();
      }
    });
  }


  const wireSizerBtn = document.getElementById("toolWireSizer");
  if (wireSizerBtn) {
    wireSizerBtn.addEventListener("click", openWireSizer);
    wireSizerBtn.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        openWireSizer();
      }
    });
  }

  const wireSizerCloseBtn = document.getElementById("wireSizerCloseBtn");
  if (wireSizerCloseBtn) {
    wireSizerCloseBtn.addEventListener("click", closeWireSizer);
  }

  const wireSizerDlg = document.getElementById("wireSizerDlg");
  if (wireSizerDlg) {
    wireSizerDlg.addEventListener("close", () => {
      const frame = document.getElementById("wireSizerFrame");
      if (frame) frame.src = "about:blank";
      setWireSizerMessage("");
    });
  }

  // Circuit Breaker AI event listeners
  const circuitBreakerBtn = document.getElementById("toolCircuitBreaker");
  if (circuitBreakerBtn) {
    circuitBreakerBtn.addEventListener("click", openCircuitBreaker);
    circuitBreakerBtn.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        openCircuitBreaker();
      }
    });
  }

  const circuitBreakerCloseBtn = document.getElementById("circuitBreakerCloseBtn");
  if (circuitBreakerCloseBtn) {
    circuitBreakerCloseBtn.addEventListener("click", closeCircuitBreaker);
  }

  const circuitBreakerDlg = document.getElementById("circuitBreakerDlg");
  if (circuitBreakerDlg) {
    circuitBreakerDlg.addEventListener("close", () => {
      if (!circuitBreakerState.running) {
        resetCircuitBreakerForm();
      }
    });
  }

  const cbBreakerDrop = document.getElementById("cbBreakerDrop");
  if (cbBreakerDrop) {
    cbBreakerDrop.addEventListener("click", () =>
      openCircuitBreakerFilePicker("breaker")
    );
    cbBreakerDrop.addEventListener("dragover", (e) =>
      handleCircuitBreakerDragOver(e, cbBreakerDrop)
    );
    cbBreakerDrop.addEventListener("dragleave", (e) =>
      handleCircuitBreakerDragLeave(e, cbBreakerDrop)
    );
    cbBreakerDrop.addEventListener("drop", (e) =>
      handleCircuitBreakerDrop("breaker", e, cbBreakerDrop)
    );
    cbBreakerDrop.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        openCircuitBreakerFilePicker("breaker");
      }
    });
  }

  const cbDirectoryDrop = document.getElementById("cbDirectoryDrop");
  if (cbDirectoryDrop) {
    cbDirectoryDrop.addEventListener("click", () =>
      openCircuitBreakerFilePicker("directory")
    );
    cbDirectoryDrop.addEventListener("dragover", (e) =>
      handleCircuitBreakerDragOver(e, cbDirectoryDrop)
    );
    cbDirectoryDrop.addEventListener("dragleave", (e) =>
      handleCircuitBreakerDragLeave(e, cbDirectoryDrop)
    );
    cbDirectoryDrop.addEventListener("drop", (e) =>
      handleCircuitBreakerDrop("directory", e, cbDirectoryDrop)
    );
    cbDirectoryDrop.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        openCircuitBreakerFilePicker("directory");
      }
    });
  }

  const cbOutputModeNew = document.getElementById("cbOutputModeNew");
  if (cbOutputModeNew) {
    cbOutputModeNew.addEventListener("click", async () => {
      circuitBreakerState.outputMode = "new";
      updateCircuitBreakerUi();
      await selectCircuitBreakerSchedulePath("new");
    });
  }

  const cbOutputModeExisting = document.getElementById("cbOutputModeExisting");
  if (cbOutputModeExisting) {
    cbOutputModeExisting.addEventListener("click", async () => {
      circuitBreakerState.outputMode = "existing";
      updateCircuitBreakerUi();
      await selectCircuitBreakerSchedulePath("existing");
    });
  }

  const cbPanelNameInput = document.getElementById("cbPanelName");
  if (cbPanelNameInput) {
    cbPanelNameInput.addEventListener("input", (e) => {
      circuitBreakerState.panelName = e.target.value;
      updateCircuitBreakerUi();
    });
  }

  const cbRunPanelBtn = document.getElementById("cbRunPanelScheduleBtn");
  if (cbRunPanelBtn) {
    cbRunPanelBtn.addEventListener("click", runCircuitBreakerInBackground);
  }

  const lightingScheduleBtn = document.getElementById("toolLightingSchedule");
  if (lightingScheduleBtn) {
    lightingScheduleBtn.addEventListener("click", openLightingSchedule);
    lightingScheduleBtn.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        openLightingSchedule();
      }
    });
  }

  const lightingScheduleCloseBtn = document.getElementById(
    "lightingScheduleCloseBtn"
  );
  if (lightingScheduleCloseBtn) {
    lightingScheduleCloseBtn.addEventListener("click", closeLightingSchedule);
  }

  const lightingScheduleDlg = document.getElementById("lightingScheduleDlg");
  if (lightingScheduleDlg) {
    lightingScheduleDlg.addEventListener("close", () => {
      save();
    });
  }

  const lightingScheduleProjectSelect = document.getElementById(
    "lightingScheduleProjectSelect"
  );
  if (lightingScheduleProjectSelect) {
    lightingScheduleProjectSelect.addEventListener("change", (e) => {
      const idx = Number(e.target.value);
      if (!Number.isNaN(idx)) setLightingScheduleProject(idx);
    });
  }

  const lightingScheduleProjectSearch = document.getElementById(
    "lightingScheduleProjectSearch"
  );
  if (lightingScheduleProjectSearch) {
    lightingScheduleProjectSearch.addEventListener("input", (e) => {
      lightingScheduleProjectQuery = e.target.value;
      renderLightingScheduleProjectOptions(lightingScheduleProjectQuery);
    });
  }

  const lightingScheduleAddRowBtn = document.getElementById(
    "lightingScheduleAddRow"
  );
  if (lightingScheduleAddRowBtn) {
    lightingScheduleAddRowBtn.addEventListener("click", addLightingScheduleRow);
  }

  const lightingTemplateSearch = document.getElementById(
    "lightingTemplateSearch"
  );
  if (lightingTemplateSearch) {
    lightingTemplateSearch.addEventListener("input", (e) => {
      lightingTemplateQuery = e.target.value;
      renderLightingTemplateList(lightingTemplateQuery);
    });
  }

  const lightingScheduleGeneralNotes = document.getElementById(
    "lightingScheduleGeneralNotes"
  );
  if (lightingScheduleGeneralNotes) {
    lightingScheduleGeneralNotes.addEventListener("input", (e) => {
      const schedule = getActiveLightingSchedule();
      if (!schedule) return;
      schedule.generalNotes = e.target.value;
      autoResizeTextarea(lightingScheduleGeneralNotes);
      debouncedSaveLightingSchedule();
    });
  }

  const lightingScheduleNotes = document.getElementById(
    "lightingScheduleNotes"
  );
  if (lightingScheduleNotes) {
    lightingScheduleNotes.addEventListener("input", (e) => {
      const schedule = getActiveLightingSchedule();
      if (!schedule) return;
      schedule.notes = e.target.value;
      autoResizeTextarea(lightingScheduleNotes);
      debouncedSaveLightingSchedule();
    });
  }






  document
    .getElementById("btnBrowseAutocad")
    .addEventListener("click", async () => {
      const res = await window.pywebview.api.select_files({
        allow_multiple: false,
        file_types: ["Executable Files (*.exe)"],
        directory: "C:\\Program Files\\Autodesk",
      });
      if (res.status === "success" && res.paths.length) {
        document.getElementById("settings_autocadPath").value = res.paths[0];
        // Uncheck radio buttons
        document
          .querySelectorAll('input[name="autocad_version_radio"]')
          .forEach((radio) => (radio.checked = false));
      }
    });

  document
    .getElementById("btnBrowseAutocadSelect")
    .addEventListener("click", async () => {
      const res = await window.pywebview.api.select_files({
        allow_multiple: false,
        file_types: ["Executable Files (*.exe)"],
        directory: "C:\\Program Files\\Autodesk",
      });
      if (res.status === "success" && res.paths.length) {
        document.getElementById("autocad_select_custom").value = res.paths[0];
        // Uncheck radio buttons
        document
          .querySelectorAll('input[name="autocad_select_radio"]')
          .forEach((radio) => (radio.checked = false));
      }
    });

  document
    .getElementById("btnSaveAutocadSelect")
    .addEventListener("click", async () => {
      const selectedRadio = document.querySelector(
        'input[name="autocad_select_radio"]:checked'
      );
      if (selectedRadio) {
        userSettings.autocadPath = selectedRadio.value;
      } else {
        userSettings.autocadPath = document
          .getElementById("autocad_select_custom")
          .value.trim();
      }
      try {
        await window.pywebview.api.save_user_settings(userSettings);
      } catch (e) {
        toast("⚠️ Could not save settings.");
      }
      closeDlg("autocadSelectDlg");
    });

  const handleAI = () => {
    if (!userSettings.apiKey) {
      toast("Setup API Key in Settings first.");
      return;
    }
    document.getElementById("emailArea").value = "";
    document.getElementById("aiSpinner").style.display = "none";
    document.getElementById("emailDlg").showModal();
  };
  document.getElementById("aiBtn").onclick = handleAI;

  document.getElementById("btnProcessEmail").onclick = async () => {
    const txt = val("emailArea");
    if (!txt) return;
    document.getElementById("aiSpinner").style.display = "flex";
    try {
      const res = await window.pywebview.api.process_email_with_ai(
        txt,
        userSettings.apiKey,
        userSettings.userName,
        userSettings.discipline
      );
      if (res.status === "success") {
        closeDlg("emailDlg");
        const aiProject = normalizeProject(res.data);
        const match = findBestProjectMatch(aiProject);
        if (match) {
          const snapshot = JSON.parse(JSON.stringify(db[match.index]));
          const target = normalizeProject(db[match.index]);
          const newDeliverable = createDeliverable({
            name: guessDeliverableName(res.data),
            due: res.data.due || "",
            notes: res.data.notes || "",
            tasks: res.data.tasks || [],
          });
          if (!target.path && aiProject.path) target.path = aiProject.path;
          if (!target.id && aiProject.id) target.id = aiProject.id;
          target.deliverables.push(newDeliverable);
          target.overviewDeliverableId = newDeliverable.id;
          db[match.index] = target;
          editIndex = match.index;
          _aiMatchSnapshot = { index: match.index, data: snapshot };
          document.getElementById("dlgTitle").textContent =
            `Edit Project \u2014 ${target.id || "Untitled"}`;
          document.getElementById("btnSaveProject").textContent = "Save Changes";
          fillForm(target);
          document.getElementById("editDlg").showModal();
          toast(
            `Matched existing project (${match.method}, ` +
            `${Math.round(match.score * 100)}% confidence).`
          );
        } else {
          openNew();
          fillForm(res.data);
        }
      } else throw new Error(res.message);
    } catch (e) {
      toast("AI Error: " + e.message);
    }
    document.getElementById("aiSpinner").style.display = "none";
  };

  document
    .getElementById("notesTextarea")
    .addEventListener("input", handleNoteInput);
  document.getElementById("settings_userName").oninput = (e) => {
    userSettings.userName = e.target.value;
    debouncedSaveUserSettings();
  };
  document.getElementById("settings_apiKey").oninput = (e) => {
    userSettings.apiKey = e.target.value;
    debouncedSaveUserSettings();
  };
  document.getElementById("settings_defaultPmInitials").oninput = (e) => {
    userSettings.defaultPmInitials = e.target.value.toUpperCase();
    debouncedSaveUserSettings();
  };
  document
    .querySelectorAll('input[name="settings_discipline_checkbox"]')
    .forEach((checkbox) => {
      checkbox.onchange = () => {
        const checked = document.querySelectorAll(
          'input[name="settings_discipline_checkbox"]:checked'
        );
        userSettings.discipline = Array.from(checked).map((cb) => cb.value);
        debouncedSaveUserSettings();
      };
    });

  ["settings_publish_autoDetectPaperSize", "publish_modal_autoDetectPaperSize"]
    .map((id) => document.getElementById(id))
    .filter(Boolean)
    .forEach((checkbox) => {
      checkbox.onchange = (e) => {
        if (!userSettings.publishDwgOptions) {
          userSettings.publishDwgOptions = { ...DEFAULT_PUBLISH_DWG_OPTIONS };
        }
        userSettings.publishDwgOptions.autoDetectPaperSize = e.target.checked;
        syncPublishOptionsInputs();
        debouncedSaveUserSettings();
      };
    });

  [
    ["settings_publish_shrinkPercent", "settings_publish_shrinkValue"],
    ["publish_modal_shrinkPercent", "publish_modal_shrinkValue"],
  ].forEach(([sliderId, labelId]) => {
    const slider = document.getElementById(sliderId);
    const label = document.getElementById(labelId);
    if (!slider) return;
    const updateLabel = () => {
      if (label) label.textContent = `${slider.value}%`;
    };
    slider.oninput = (e) => {
      if (!userSettings.publishDwgOptions) {
        userSettings.publishDwgOptions = { ...DEFAULT_PUBLISH_DWG_OPTIONS };
      }
      const nextValue = Number(e.target.value);
      userSettings.publishDwgOptions.shrinkPercent = Number.isFinite(nextValue)
        ? nextValue
        : 100;
      updateLabel();
      syncPublishOptionsInputs();
      debouncedSaveUserSettings();
    };
    updateLabel();
  });

  ["settings_freeze_scanAllLayers", "freeze_modal_scanAllLayers"]
    .map((id) => document.getElementById(id))
    .filter(Boolean)
    .forEach((checkbox) => {
      checkbox.onchange = (e) => {
        if (!userSettings.freezeLayerOptions) {
          userSettings.freezeLayerOptions = { ...DEFAULT_FREEZE_LAYER_OPTIONS };
        }
        userSettings.freezeLayerOptions.scanAllLayers = e.target.checked;
        syncFreezeOptionsInputs();
        debouncedSaveUserSettings();
      };
    });

  ["settings_thaw_scanAllLayers", "thaw_modal_scanAllLayers"]
    .map((id) => document.getElementById(id))
    .filter(Boolean)
    .forEach((checkbox) => {
      checkbox.onchange = (e) => {
        if (!userSettings.thawLayerOptions) {
          userSettings.thawLayerOptions = { ...DEFAULT_THAW_LAYER_OPTIONS };
        }
        userSettings.thawLayerOptions.scanAllLayers = e.target.checked;
        syncThawOptionsInputs();
        debouncedSaveUserSettings();
      };
    });

  const cleanOptionBindings = [
    ["settings_clean_stripXrefs", "stripXrefs"],
    ["settings_clean_setByLayer", "setByLayer"],
    ["settings_clean_purge", "purge"],
    ["settings_clean_audit", "audit"],
    ["settings_clean_hatchColor", "hatchColor"],
    ["clean_modal_stripXrefs", "stripXrefs"],
    ["clean_modal_setByLayer", "setByLayer"],
    ["clean_modal_purge", "purge"],
    ["clean_modal_audit", "audit"],
    ["clean_modal_hatchColor", "hatchColor"],
  ];
  cleanOptionBindings.forEach(([id, key]) => {
    const checkbox = document.getElementById(id);
    if (!checkbox) return;
    checkbox.onchange = (e) => {
      if (!userSettings.cleanDwgOptions) {
        userSettings.cleanDwgOptions = { ...DEFAULT_CLEAN_DWG_OPTIONS };
      }
      userSettings.cleanDwgOptions[key] = e.target.checked;
      syncCleanOptionsInputs();
      debouncedSaveUserSettings();
    };
  });

  document.getElementById("btnPasteImport").onclick = () => {
    const rows = parseCSV(val("pasteArea"));
    importRows(rows, document.getElementById("hasHeader").checked);
    closeDlg("pasteDlg");
  };

  document.getElementById("btnMarkOverdue").onclick = () => {
    document.getElementById("markOverdueDlg").showModal();
  };
  document.getElementById("btnConfirmMarkOverdue").onclick = async () => {
    try {
      const response =
        await window.pywebview.api.mark_overdue_projects_complete();
      if (response.status === "success") {
        toast(`Marked ${response.count} deliverables as complete.`);
        db = await load();
        render();
      } else {
        toast("Failed to mark deliverables as complete.");
      }
    } catch (e) {
      toast("Error marking deliverables as complete.");
    }
    closeDlg("markOverdueDlg");
  };
  document.getElementById("btnMarkOverdueDelivered").onclick = () => {
    document.getElementById("markOverdueDeliveredDlg").showModal();
  };
  document.getElementById("btnConfirmMarkOverdueDelivered").onclick =
    async () => {
      try {
        const response =
          await window.pywebview.api.mark_overdue_projects_delivered();
        if (response.status === "success") {
          toast(`Marked ${response.count} deliverables as delivered.`);
          db = await load();
          render();
        } else {
          toast("Failed to mark deliverables as delivered.");
        }
      } catch (e) {
        toast("Error marking deliverables as delivered.");
      }
      closeDlg("markOverdueDeliveredDlg");
    };

  document.getElementById("btnDeleteAll").onclick = () => {
    document.getElementById("deleteConfirmInput").value = "";
    document.getElementById("btnDeleteConfirm").disabled = true;
    document.getElementById("deleteDlg").showModal();
  };
  document.getElementById("deleteConfirmInput").oninput = (e) => {
    document.getElementById("btnDeleteConfirm").disabled =
      e.target.value !== "DELETE";
  };
  document.getElementById("btnDeleteConfirm").onclick = () => {
    db = [];
    save();
    render();
    closeDlg("deleteDlg");
  };

  document.getElementById("btnDeleteAllNotes").onclick = () => {
    document.getElementById("deleteNotesConfirmInput").value = "";
    document.getElementById("btnDeleteNotesConfirm").disabled = true;
    document.getElementById("deleteNotesDlg").showModal();
  };
  document.getElementById("deleteNotesConfirmInput").oninput = (e) => {
    document.getElementById("btnDeleteNotesConfirm").disabled =
      e.target.value !== "DELETE";
  };
  document.getElementById("btnDeleteNotesConfirm").onclick = async () => {
    try {
      const response = await window.pywebview.api.delete_all_notes();
      if (response.status === "success") {
        toast("All notes data deleted.");
        notesDb = {};
        noteTabs = ["General"];
        activeNoteTab = "General";
        renderNoteTabs();
      } else {
        toast("Failed to delete notes data.");
      }
    } catch (e) {
      toast("Error deleting notes data.");
    }
    closeDlg("deleteNotesDlg");
  };

  document.getElementById("btnUninstallAllPlugins").onclick = async () => {
    try {
      const response = await window.pywebview.api.check_autocad_running();
      if (response.is_running) {
        toast(
          "AutoCAD is currently running. Please close AutoCAD and try again."
        );
        return;
      }
      document.getElementById("uninstallPluginsConfirmInput").value = "";
      document.getElementById("btnUninstallPluginsConfirm").disabled = true;
      document.getElementById("uninstallPluginsDlg").showModal();
    } catch (e) {
      toast("Error checking AutoCAD status.");
    }
  };
  document.getElementById("uninstallPluginsConfirmInput").oninput = (e) => {
    document.getElementById("btnUninstallPluginsConfirm").disabled =
      e.target.value !== "UNINSTALL";
  };
  document.getElementById("btnUninstallPluginsConfirm").onclick = async () => {
    try {
      const response = await window.pywebview.api.uninstall_all_plugins();
      if (response.status === "success") {
        toast(`Uninstalled ${response.count} plugins.`);
        ensureBundlesRendered({ force: true });
      } else {
        toast("Failed to uninstall plugins.");
      }
    } catch (e) {
      toast("Error uninstalling plugins.");
    }
    closeDlg("uninstallPluginsDlg");
  };

  document
    .getElementById("commands-container")
    .addEventListener("click", handleBundleAction);
  document.getElementById("btnInstallPrereq").onclick = async () => {
    const spinner = document.getElementById("installSpinner");
    spinner.style.display = "block";
    try {
      const all = await window.pywebview.api.get_bundle_statuses();
      const bundle = all.data.find((b) => b.name.includes("CleanCADCommands"));
      if (bundle) {
        await window.pywebview.api.install_single_bundle(bundle.asset);
        toast("Installed!");
        closeDlg("installPrereqDlg");
        updateCleanDwgToolState();
        ensureBundlesRendered({ force: true });
      }
    } catch (e) {
      toast("Installation failed.");
    }
    spinner.style.display = "none";
  };
}

async function init() {
  showAppLoader();
  try {
    if (!window.pywebview)
      await new Promise((r) => window.addEventListener("pywebviewready", r));
    initEventListeners();
    initTabbedInterfaces();
    updateStickyOffsets();
    refreshAppUpdateStatus();
    await loadUserSettings();
    initThemeFromPreferences();

    if (isNewUser()) {
      hideMainApp();
      showOnboardingModal();
    } else {
      showMainApp();
      const [loadedDb, loadedNotes, loadedTimesheets, loadedTemplates] = await Promise.all([
        load(),
        loadNotes(),
        loadTimesheets(),
        loadTemplates(),
      ]);
      db = loadedDb;
      notesDb = loadedNotes || {};
      timesheetDb = loadedTimesheets || { weeks: {}, lastModified: null };
      templatesDb = loadedTemplates || { templates: [], defaultTemplatesInstalled: false, lastModified: null };
      renderNoteTabs();
      render();

      // Show setup help banner for returning users who haven't disabled it
      if (userSettings.showSetupHelp !== false) {
        // Delay showing the banner slightly so the UI is fully loaded
        setTimeout(() => showSetupHelpBanner(), 1000);
      }
    }
    prefetchBundles();
  } finally {
    hideAppLoader();
  }
}

init();
