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
  checklists:
    "https://brainy-seahorse-3c5.notion.site/Checklists-2b13fdbb662c80dd86c8e9d3f0d65123",
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
const LIGHTING_SCHEDULE_SYNC_SCHEMA_VERSION = "1.0.0";
const LIGHTING_SCHEDULE_SYNC_FILE_NAME = "T24LightingFixtureSchedule.sync.json";
const TITLE24_SCHEMA_VERSION = "1.0.0";
const TITLE24_SCOPE_OPTIONS_FILE_PATH =
  "assets/title24/energycodeace.indoor.page1.options.json";
const TITLE24_PAGE_COUNT = 4;
const TITLE24_SCOPE_OPTION_FIELDS = [
  "occupancyType",
  "projectScopeType",
  "lightingSystemType",
];
const STAR_ICON_PATH =
  "M12 17.27L18.18 21l-1.64-7.03L22 9.24l-7.19-.61L12 2 9.19 8.63 2 9.24l5.46 4.73L5.82 21z";
const EYE_ICON_PATH =
  "M12 4.5C7 4.5 2.73 7.61 1 12c1.73 4.39 6 7.5 11 7.5s9.27-3.11 11-7.5C21.27 7.61 17 4.5 12 4.5zm0 12.5a5 5 0 1 1 0-10 5 5 0 0 1 0 10zm0-8a3 3 0 1 0 0 6 3 3 0 0 0 0-6z";
const PIN_ICON_PATH =
  "M12 2C8.13 2 5 5.13 5 9c0 2.76 1.87 5.08 4.42 5.76L10 22l2-3 2 3-.42-7.24C16.13 14.08 18 11.76 18 9c0-3.87-3.13-7-7-7zm0 8.5A2.5 2.5 0 1 1 12 5a2.5 2.5 0 0 1 0 5z";
const PENCIL_ICON_PATH =
  "M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zm17.71-10.21c.39-.39.39-1.02 0-1.41l-2.34-2.34a1 1 0 0 0-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z";
const TRASH_ICON_PATH =
  "M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zm13-15h-3.5l-1-1h-5l-1 1H5v2h14V4z";
const MAIL_ICON_PATH =
  "M20 4H4c-1.1 0-2 .9-2 2v12a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V6c0-1.1-.9-2-2-2zm-.8 2L12 11.2 4.8 6h14.4zM4 18V7l7.4 5.2a1 1 0 0 0 1.2 0L20 7v11H4z";

// Checklist Icons
const CHECKLIST_ICON_PATH =
  "M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V5h14v14zM7 10h2v7H7zm4-3h2v10h-2zm4 6h2v4h-2z";
const WORKROOM_ICON_PATH =
  "M3 5c0-1.1.9-2 2-2h6v4h2V3h6c1.1 0 2 .9 2 2v4H3V5zm0 6h18v8c0 1.1-.9 2-2 2H5c-1.1 0-2-.9-2-2v-8zm7 2v2h4v-2h-4z";

const CHECK_ICON_PATH =
  "M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z";
const MAX_DELIVERABLE_EMAIL_REFS = 3;

// Checklists Data
let checklistsDb = {
  checklists: [],
  lastModified: null,
};

// Active checklist state for modal
let activeChecklistProject = null;
let activeChecklistDeliverable = null;
let activeChecklistTab = null;
let activeWorkroomLeftTab = "tasks";
let workroomToolStatusState = {
  toolId: "",
  label: "",
  message: "Ready.",
  phase: "idle",
};

// ===================== CHECKLISTS SYSTEM =====================

function generateChecklistId() {
  return "checklist_" + Math.random().toString(36).substr(2, 9);
}

function generateChecklistItemId() {
  return "item_" + Math.random().toString(36).substr(2, 9);
}

function generateChecklistInstanceId() {
  return "instance_" + Math.random().toString(36).substr(2, 9);
}

// Default checklist factory - Electrical Plan Check Checklist
function getDefaultChecklist() {
  return {
    id: "checklist_default",
    name: "Default",
    isDefault: true,
    isLocked: true,
    items: [
      { id: generateChecklistItemId(), text: "Check relevant codes that apply to the project based on local city, state, and national codes. (CEC 90.2, 90.4)", order: 0, isDefault: true },
      { id: generateChecklistItemId(), text: "Check if the tenant space has existing mechanical units on the roof that are not powered from tenant electrical panels. (CEC 430.102, 440.14)", order: 1, isDefault: true },
      { id: generateChecklistItemId(), text: "Check for unmentioned items that could need power (interior signage). (CEC 600.6)", order: 2, isDefault: true },
      { id: generateChecklistItemId(), text: "Check for occ sensor on ceiling instead of wall in storage room. (Title 24, Part 6 Section 130.1(c)1)", order: 3, isDefault: true },
      { id: generateChecklistItemId(), text: "Check for dedicated service receptacle within 25ft of electrical panels. (CEC 210.63(B)(2), 110.26(E))", order: 4, isDefault: true },
      { id: generateChecklistItemId(), text: "Check for junction box indicated as wall mount for hand dryers. (CEC 314.23, 314.29)", order: 5, isDefault: true },
      { id: generateChecklistItemId(), text: "Check for adequate space for electrical panels, relocate to BOH corridors as necessary, avoid storage rooms, IT server racks. (CEC 110.26(A), 110.26(E))", order: 6, isDefault: true },
      { id: generateChecklistItemId(), text: "Check for food waste disposer for all sinks with outlet under counter and switch above (confirm with plumbing as it is not a requirement). (CEC 422.16(B)(1), 422.31(B))", order: 7, isDefault: true },
      { id: generateChecklistItemId(), text: "Check for furniture systems and include note to verify point of connection for furniture systems. (CEC 605)", order: 8, isDefault: true },
      { id: generateChecklistItemId(), text: "Check for controlled receptacles in office, lobby, kitchen, printer/copy room, conference room, meeting room. Modular furniture workstations need at least one controlled receptacle per workstation. (Title 24, Part 6 Section 130.5(d))", order: 9, isDefault: true },
      { id: generateChecklistItemId(), text: "Check for tamper proof receptacles in areas where children may be present: business offices, lobbies, waiting areas, theaters, auditoriums, gyms, bowling alleys, bus stations, airports, train stations. (CEC 406.12)", order: 10, isDefault: true },
      { id: generateChecklistItemId(), text: "Check for rooftop mechanical units shown on RCP, ensure they are dashed in appearance and noted to go on the roof along with rooftop receptacle. (CEC 210.63(A), 440.14)", order: 11, isDefault: true },
      { id: generateChecklistItemId(), text: "Check for return or supply air system over 2000CFM for duct smoke requirement. Any mechanical units 2000CFM or over should get duct smoke. (Title 24, Part 2 CBC Section 907.2.12.1.2)", order: 12, isDefault: true },
      { id: generateChecklistItemId(), text: "Check for hand dryer specification, if not add note to confirm exact breaker size with manufacturer. (CEC 110.3(B))", order: 13, isDefault: true },
      { id: generateChecklistItemId(), text: "Check for kAIC rating shown on single line main switchboard. (CEC 110.9, 110.10)", order: 14, isDefault: true },
      { id: generateChecklistItemId(), text: "Check for >4000W in space needing demand response. Provide necessary software and device(s) to automatically reducing the lighting power by at least 15% upon receiving a demand response signal. (Title 24, Part 6 Section 110.12)", order: 15, isDefault: true },
      { id: generateChecklistItemId(), text: "Check for GFCI protection at required nonresidential receptacle locations (kitchens, outdoor, rooftops, within 6 ft of sinks, etc.). (CEC 210.8(B))", order: 16, isDefault: true },
      { id: generateChecklistItemId(), text: "Check meeting rooms for required receptacle outlets, including floor outlets when room size thresholds are met. (CEC 210.65)", order: 17, isDefault: true },
      { id: generateChecklistItemId(), text: "Check voltage drop requirements and calculations for feeders and branch circuits (<=5% combined). (Title 24, Part 6 Section 130.5(c))", order: 18, isDefault: true },
      { id: generateChecklistItemId(), text: "Check AIC ratings for panels and transformers on single line. (CEC 110.9, 110.10)", order: 19, isDefault: true },
      { id: generateChecklistItemId(), text: "Check for daylight harvesting in daylit zones. Provide daylight sensor & room controller for automatic dimming of light fixtures. (Title 24, Part 6 Section 130.1(d))", order: 20, isDefault: true },
      { id: generateChecklistItemId(), text: "Check for dimming in all rooms that are BOTH >100sqft AND >0.5W/sqft.", order: 21, isDefault: true },
      { id: generateChecklistItemId(), text: "Check for occupancy sensors which are required in offices, conference & meeting rooms, classrooms, restrooms, multipurpose rooms, warehouses, library aisles, corridors & stairwells.", order: 22, isDefault: true },
      { id: generateChecklistItemId(), text: "Check for time-switch controls which are allowed in lobbies, retail sales floors, commercial kitchens, auditoriums & theaters, and large multipurpose rooms.", order: 23, isDefault: true },
      { id: generateChecklistItemId(), text: "Check for 2-hour bypass switch in each room controlled by automatic time controlled on/off. Provide 1 bypass switch per 5000sqft of room size.", order: 24, isDefault: true },
      { id: generateChecklistItemId(), text: "Check for timeclock to indoor light fixtures that has manual override up to 2-hours, battery or internal memory capable of storing schedule for at least 7 days if power goes out.", order: 25, isDefault: true },
      { id: generateChecklistItemId(), text: "Check for astronomical timeclock, photosensor located on the roof, and contactors with signage, exterior / outdoor lights, pole lights, etc. (Section 130.2(c)2.B)", order: 26, isDefault: true },
      { id: generateChecklistItemId(), text: "Check for spaces that don't require multilevel controls such as: rooms under 100sqft, restrooms, rooms with only one luminaire.", order: 27, isDefault: true }
    ]
  };
}

async function loadChecklists() {
  try {
    const data = await window.pywebview.api.get_checklists();
    checklistsDb = data || { checklists: [], lastModified: null };

    // Ensure default checklist exists
    if (!checklistsDb.checklists.find(c => c.isDefault)) {
      checklistsDb.checklists.unshift(getDefaultChecklist());
    }

    return checklistsDb;
  } catch (e) {
    console.warn("Failed to load checklists:", e);
    // Initialize with default if load fails
    checklistsDb = { checklists: [getDefaultChecklist()], lastModified: null };
    return checklistsDb;
  }
}

async function saveChecklists() {
  try {
    checklistsDb.lastModified = new Date().toISOString();
    const response = await window.pywebview.api.save_checklists(checklistsDb);
    if (response.status !== "success") throw new Error(response.message);
  } catch (e) {
    console.warn("Failed to save checklists:", e);
    toast("Failed to save checklists.");
  }
}

function getChecklistById(id) {
  return checklistsDb.checklists.find(c => c.id === id);
}

function createChecklist(name) {
  const newChecklist = {
    id: generateChecklistId(),
    name: name,
    isDefault: false,
    isLocked: false,
    items: []
  };
  checklistsDb.checklists.push(newChecklist);
  saveChecklists();
  return newChecklist;
}

function deleteChecklist(id) {
  const idx = checklistsDb.checklists.findIndex(c => c.id === id);
  if (idx > -1 && !checklistsDb.checklists[idx].isLocked) {
    checklistsDb.checklists.splice(idx, 1);
    saveChecklists();
    return true;
  }
  return false;
}

function addChecklistItem(checklistId, text) {
  const checklist = getChecklistById(checklistId);
  if (checklist) {
    const newItem = {
      id: generateChecklistItemId(),
      text: text,
      order: checklist.items.length,
      isDefault: false
    };
    checklist.items.push(newItem);
    saveChecklists();
    return newItem;
  }
  return null;
}

function removeChecklistItem(checklistId, itemId) {
  const checklist = getChecklistById(checklistId);
  if (checklist) {
    const idx = checklist.items.findIndex(i => i.id === itemId);
    if (idx > -1) {
      checklist.items.splice(idx, 1);
      // Reorder remaining items
      checklist.items.forEach((item, i) => item.order = i);
      saveChecklists();
      return true;
    }
  }
  return false;
}

function updateChecklistItem(checklistId, itemId, text) {
  const checklist = getChecklistById(checklistId);
  if (checklist) {
    const item = checklist.items.find(i => i.id === itemId);
    if (item) {
      item.text = text;
      saveChecklists();
      return true;
    }
  }
  return false;
}

function resetChecklistToDefault(checklistId) {
  const checklist = getChecklistById(checklistId);
  if (checklist && checklist.isDefault) {
    const defaultChecklist = getDefaultChecklist();
    checklist.items = defaultChecklist.items.map(item => ({ ...item }));
    saveChecklists();
    return true;
  }
  return false;
}

// Timesheet Constants
const DISCIPLINE_TO_FUNCTION = {
  Electrical: "E",
  Mechanical: "M",
  Plumbing: "P",
};
const WORKROOM_CAD_DISCIPLINES = ["Electrical", "Mechanical", "Plumbing"];
const WORKROOM_CAD_TOOL_IDS = new Set([
  "toolPublishDwgs",
  "toolFreezeLayers",
  "toolThawLayers",
  "toolCleanXrefs",
]);
const WORKROOM_TEMPLATE_TOOL_IDS = new Set([
  "toolCreateNarrativeTemplate",
  "toolCreatePlanCheckTemplate",
]);
const WORKROOM_LAUNCH_CONTEXT_TOOL_IDS = new Set([
  ...WORKROOM_CAD_TOOL_IDS,
  ...WORKROOM_TEMPLATE_TOOL_IDS,
]);
const WORKROOM_HIDDEN_TOOL_IDS = new Set([
  "toolLightingSchedule",
  "toolTitle24Compliance",
]);
const WORKROOM_DISCIPLINE_KEYWORDS = {
  Electrical: [
    /\belectrical\b/gi,
    /\bpanel\b/gi,
    /\blighting\b/gi,
    /\bcircuit\b/gi,
    /\breceptacle\b/gi,
    /\bswitchgear\b/gi,
    /\btransformer\b/gi,
    /\bbreaker\b/gi,
  ],
  Mechanical: [
    /\bmechanical\b/gi,
    /\bhvac\b/gi,
    /\ba\/c\b/gi,
    /\bac\b/gi,
    /\brtu\b/gi,
    /\bahu\b/gi,
    /\bair handler\b/gi,
    /\bduct\b/gi,
    /\bsupply air\b/gi,
    /\breturn air\b/gi,
    /\bexhaust\b/gi,
  ],
  Plumbing: [
    /\bplumbing\b/gi,
    /\bplumb\b/gi,
    /\bpipe\b/gi,
    /\bpiping\b/gi,
    /\bwater\b/gi,
    /\bsanitary\b/gi,
    /\bsewer\b/gi,
    /\bwaste\b/gi,
    /\bdrain\b/gi,
    /\bvent\b/gi,
  ],
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

function highlightText(text, query) {
  const frag = document.createDocumentFragment();
  if (!query || !text) {
    frag.appendChild(document.createTextNode(text || ""));
    return frag;
  }
  const lower = text.toLowerCase();
  const qLen = query.length;
  let lastIdx = 0;
  let idx = lower.indexOf(query, lastIdx);
  while (idx !== -1) {
    if (idx > lastIdx) {
      frag.appendChild(document.createTextNode(text.slice(lastIdx, idx)));
    }
    const mark = document.createElement("mark");
    mark.className = "search-highlight";
    mark.textContent = text.slice(idx, idx + qLen);
    frag.appendChild(mark);
    lastIdx = idx + qLen;
    idx = lower.indexOf(query, lastIdx);
  }
  if (lastIdx < text.length) {
    frag.appendChild(document.createTextNode(text.slice(lastIdx)));
  }
  return frag;
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

function getWorkroomAvailableDisciplines() {
  const disciplines = normalizeDisciplineList(userSettings.discipline).filter((discipline) =>
    WORKROOM_CAD_DISCIPLINES.includes(discipline)
  );
  return disciplines.length ? disciplines : ["Electrical"];
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

function getTemplateToolIdForKey(templateKey) {
  if (templateKey === "narrative") return "toolCreateNarrativeTemplate";
  if (templateKey === "planCheck") return "toolCreatePlanCheckTemplate";
  return "";
}

function clearTemplateToolRunState(toolId) {
  if (!toolId) return;
  const card = document.getElementById(toolId);
  if (!card) return;
  card.classList.remove("running");
  const statusEl = card.querySelector(".tool-card-status");
  if (statusEl) statusEl.textContent = "";
  const checklistModal = document.getElementById("checklistModal");
  if (checklistModal?.open) resetWorkroomToolStatus();
}

function getTemplateToolContext(options = {}) {
  const launchContext = options?.launchContext || null;
  const launchSource = String(launchContext?.source || "").trim().toLowerCase();
  if (launchSource === "workroom") {
    const { project, deliverable } = getActiveWorkroomContext();
    return {
      projectPath: String(launchContext?.projectPath || project?.path || "").trim(),
      projectName: String(
        project?.name || project?.nick || project?.id || ""
      ).trim(),
      deliverableName: String(deliverable?.name || "").trim(),
      source: "workroom",
    };
  }

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

  return { projectPath, projectName, deliverableName: "", source: "default" };
}

async function handleTemplateToolSave(templateKey, label, options = {}) {
  const launchContext = options?.launchContext || null;
  const toolId = String(
    options?.toolId || getTemplateToolIdForKey(templateKey)
  ).trim();
  const card = toolId ? document.getElementById(toolId) : null;
  if (card?.classList.contains("running")) return;
  if (card) {
    card.classList.add("running");
    window.updateToolStatus(toolId, "Initializing...");
  }

  try {
    await loadTemplates();
  } catch (e) {
    console.warn("Failed to refresh templates:", e);
  }

  const template = getTemplateByKey(templateKey);
  if (!template) {
    if (toolId) {
      window.updateToolStatus(
        toolId,
        `ERROR: Template "${label}" not found.`
      );
    }
    toast(`Template "${label}" not found.`);
    return;
  }

  const { projectPath, projectName, deliverableName } = getTemplateToolContext({
    launchContext,
  });
  const context = {};
  if (projectName) context.projectName = projectName;
  if (deliverableName) context.deliverableName = deliverableName;
  let manualDefaultDir = projectPath || null;

  const isWorkroomLaunch =
    String(launchContext?.source || "").trim().toLowerCase() === "workroom";
  if (isWorkroomLaunch && window.pywebview?.api?.create_template_for_workroom) {
    try {
      const autoResult = await window.pywebview.api.create_template_for_workroom(
        templateKey,
        launchContext,
        context,
        "timestamp"
      );
      if (autoResult?.status === "success" && autoResult.path) {
        if (toolId) window.updateToolStatus(toolId, "DONE");
        toast("Template saved.");
        return autoResult;
      }
      if (autoResult?.status && autoResult.status !== "fallback") {
        const message =
          autoResult?.message || "Failed to create template automatically.";
        if (toolId) window.updateToolStatus(toolId, `ERROR: ${message}`);
        toast(message);
        return autoResult;
      }
      if (autoResult?.status === "fallback") {
        manualDefaultDir = null;
        console.info(
          `Workroom template auto-create fallback (${templateKey}):`,
          autoResult?.message || "manual save dialog"
        );
      }
    } catch (e) {
      const message = e?.message || "Failed to create template automatically.";
      if (toolId) window.updateToolStatus(toolId, `ERROR: ${message}`);
      toast(message);
      return;
    }
  }

  const defaultName = String(label || template.name || "Template").trim() || "Template";
  let selection = null;
  try {
    selection = await window.pywebview.api.select_template_save_location(
      manualDefaultDir,
      defaultName,
      template.fileType
    );
  } catch (e) {
    if (toolId) window.updateToolStatus(toolId, "ERROR: Error selecting save location.");
    toast("Error selecting save location.");
    return;
  }

  if (!selection || selection.status === "cancelled") {
    clearTemplateToolRunState(toolId);
    return;
  }
  if (selection.status !== "success" || !selection.path) {
    if (toolId) {
      window.updateToolStatus(
        toolId,
        `ERROR: ${selection.message || "Failed to select save location."}`
      );
    }
    toast(selection.message || "Failed to select save location.");
    return;
  }

  const templateOptions = { templateKey };

  try {
    const result = await window.pywebview.api.copy_template_to_path(
      template.id,
      selection.path,
      context,
      templateOptions
    );
    if (result.status === "success") {
      if (toolId) window.updateToolStatus(toolId, "DONE");
      toast("Template saved.");
    } else {
      if (toolId) {
        window.updateToolStatus(
          toolId,
          `ERROR: ${result.message || "Failed to save template."}`
        );
      }
      toast(result.message || "Failed to save template.");
    }
  } catch (e) {
    if (toolId) window.updateToolStatus(toolId, "ERROR: Error saving template.");
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
  if (id === "editDlg") {
    flushModalEmailSession(false);
  }
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
function positionInlineCalendarPopup(anchorElement, calendarContainer, anchorParent) {
  const gap = 6;
  const rect = anchorElement.getBoundingClientRect();
  const isBodyParent = anchorParent === document.body;
  const scrollTop = isBodyParent
    ? (window.scrollY || document.documentElement.scrollTop || 0)
    : anchorParent.scrollTop;
  const scrollLeft = isBodyParent
    ? (window.scrollX || document.documentElement.scrollLeft || 0)
    : anchorParent.scrollLeft;
  const viewHeight = isBodyParent ? window.innerHeight : anchorParent.clientHeight;
  const viewWidth = isBodyParent ? window.innerWidth : anchorParent.clientWidth;

  calendarContainer.style.position = "absolute";
  calendarContainer.style.zIndex = "999999";
  calendarContainer.style.top = "0px";
  calendarContainer.style.left = "0px";
  calendarContainer.style.visibility = "hidden";

  const popupHeight = calendarContainer.offsetHeight;
  const popupWidth = calendarContainer.offsetWidth;
  let anchorTop;
  let anchorBottom;
  let anchorLeft;
  let viewTop;
  let viewBottom;
  let viewLeft;
  let viewRight;

  if (isBodyParent) {
    // Body coordinates are page-based; using parentRect here would double-count scroll.
    anchorTop = rect.top + scrollTop;
    anchorBottom = rect.bottom + scrollTop;
    anchorLeft = rect.left + scrollLeft;
    viewTop = scrollTop;
    viewBottom = scrollTop + viewHeight;
    viewLeft = scrollLeft;
    viewRight = scrollLeft + viewWidth;
  } else {
    const parentRect = anchorParent.getBoundingClientRect();
    anchorTop = rect.top - parentRect.top + scrollTop;
    anchorBottom = rect.bottom - parentRect.top + scrollTop;
    anchorLeft = rect.left - parentRect.left + scrollLeft;
    viewTop = scrollTop;
    viewBottom = scrollTop + viewHeight;
    viewLeft = scrollLeft;
    viewRight = scrollLeft + viewWidth;
  }

  const availableAbove = anchorTop - viewTop;
  const availableBelow = viewBottom - anchorBottom;

  let top;
  if (availableBelow >= popupHeight + gap) {
    top = anchorBottom + gap;
  } else if (availableAbove >= popupHeight + gap) {
    top = anchorTop - popupHeight - gap;
  } else if (availableBelow >= availableAbove) {
    top = Math.max(viewTop + gap, viewBottom - popupHeight - gap);
  } else {
    top = Math.max(viewTop + gap, anchorTop - popupHeight - gap);
  }

  let left = anchorLeft;
  const maxLeft = Math.max(viewLeft + gap, viewRight - popupWidth - gap);
  left = Math.min(Math.max(left, viewLeft + gap), maxLeft);

  calendarContainer.style.top = `${top}px`;
  calendarContainer.style.left = `${left}px`;
  calendarContainer.style.visibility = "";
}

function showCalendarForInput(inputElement, onSelectCallback) {
  // Remove any existing calendar
  const existingCalendar = document.getElementById('inlineCalendarPicker');
  if (existingCalendar) existingCalendar.remove();

  // Always open to the current month/day.
  const openDate = new Date();

  // Create calendar container
  const calendarContainer = el('div', {
    className: 'inline-calendar-picker',
    id: 'inlineCalendarPicker'
  });

  // Render calendar
  const calendar = renderInlineCalendar(
    openDate.getFullYear(),
    openDate.getMonth(),
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
  positionInlineCalendarPopup(inputElement, calendarContainer, anchorParent);

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

  // Always open to the current month/day.
  const openDate = new Date();

  // Create calendar container
  const calendarContainer = el('div', {
    className: 'inline-calendar-picker',
    id: 'inlineCalendarPicker'
  });

  // Render calendar
  const calendar = renderInlineCalendar(
    openDate.getFullYear(),
    openDate.getMonth(),
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
  positionInlineCalendarPopup(anchorElement, calendarContainer, anchorParent);

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

function fileUrlToLocalPath(url) {
  try {
    const parsed = new URL(url);
    if (parsed.protocol !== "file:") return "";
    let path = decodeURIComponent(parsed.pathname || "");
    if (/^\/[A-Za-z]:/.test(path)) path = path.slice(1);
    if (parsed.host) {
      const uncPath = path.replace(/\//g, "\\");
      return `\\\\${parsed.host}${uncPath}`;
    }
    return path.replace(/\//g, "\\");
  } catch {
    return "";
  }
}

function isEmailFilePath(value) {
  const s = String(value || "").trim();
  return /\.(msg|eml)$/i.test(s);
}

function inferEmailRefSource(raw = "", url = "") {
  const joined = `${raw} ${url}`.toLowerCase();
  if (joined.includes(".msg") || joined.includes(".eml") || /^file:\/\//i.test(url)) {
    return "file";
  }
  if (/^outlook:/i.test(raw) || /^outlook:/i.test(url)) return "outlook-url";
  if (/^mailto:/i.test(raw) || /^mailto:/i.test(url)) return "mailto";
  return "url";
}

function normalizeEmailRef(value) {
  if (!value) return null;
  if (typeof value === "string") {
    const link = normalizeLink(value);
    if (!link.raw) return null;
    return {
      raw: link.raw,
      url: link.url || link.raw,
      label: link.label || "Email",
      source: inferEmailRefSource(link.raw, link.url),
      savedAt: "",
    };
  }
  if (typeof value !== "object") return null;
  const raw = String(value.raw || value.url || "").trim();
  if (!raw) return null;
  const link = normalizeLink(raw);
  const url = String(value.url || link.url || raw).trim();
  const label = String(value.label || link.label || "Email").trim() || "Email";
  const source = String(value.source || inferEmailRefSource(raw, url)).trim();
  const savedAt = String(value.savedAt || value.saved_at || "").trim();
  return { raw, url, label, source, savedAt };
}

function normalizeEmailRefKey(value) {
  const ref = normalizeEmailRef(value);
  if (!ref) return "";
  return [
    String(ref.raw || "").toLowerCase(),
    String(ref.url || "").toLowerCase(),
    String(ref.source || "").toLowerCase(),
  ].join("|");
}

function sameEmailRef(a, b) {
  const left = normalizeEmailRef(a);
  const right = normalizeEmailRef(b);
  if (!left && !right) return true;
  if (!left || !right) return false;
  return (
    String(left.raw || "").toLowerCase() === String(right.raw || "").toLowerCase() &&
    String(left.url || "").toLowerCase() === String(right.url || "").toLowerCase() &&
    String(left.source || "").toLowerCase() === String(right.source || "").toLowerCase()
  );
}

function normalizeEmailRefs(value, fallbackSingle = null) {
  const candidates = [];
  if (Array.isArray(value)) {
    candidates.push(...value);
  } else if (value != null) {
    candidates.push(value);
  }
  if (fallbackSingle != null) candidates.push(fallbackSingle);

  const seen = new Set();
  const out = [];
  for (const candidate of candidates) {
    const normalized = normalizeEmailRef(candidate);
    if (!normalized) continue;
    const key = normalizeEmailRefKey(normalized);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    out.push(normalized);
    if (out.length >= MAX_DELIVERABLE_EMAIL_REFS) break;
  }
  return out;
}

function sameEmailRefList(a, b) {
  const left = normalizeEmailRefs(a);
  const right = normalizeEmailRefs(b);
  if (left.length !== right.length) return false;
  for (let i = 0; i < left.length; i++) {
    if (!sameEmailRef(left[i], right[i])) return false;
  }
  return true;
}

function syncDeliverableEmailRefs(deliverable) {
  if (!deliverable || typeof deliverable !== "object") return [];
  const refs = normalizeEmailRefs(deliverable.emailRefs, deliverable.emailRef);
  deliverable.emailRefs = refs;
  deliverable.emailRef = refs[0] || null;
  return refs;
}

function getEmailRefLocalPath(emailRef) {
  const ref = normalizeEmailRef(emailRef);
  if (!ref) return "";
  const raw = String(ref.raw || "").trim();
  if (/^[A-Za-z]:[\\/]/.test(raw) || /^\\\\/.test(raw)) return raw;
  if (/^file:\/\//i.test(raw)) return fileUrlToLocalPath(raw);
  const url = String(ref.url || "").trim();
  if (/^file:\/\//i.test(url)) return fileUrlToLocalPath(url);
  return "";
}

function isManagedSavedEmailRef(emailRef) {
  const ref = normalizeEmailRef(emailRef);
  return !!ref && String(ref.source || "").toLowerCase() === "saved-file";
}

async function deleteManagedSavedEmailRef(emailRef) {
  const path = getEmailRefLocalPath(emailRef);
  if (!path || !isManagedSavedEmailRef(emailRef)) return;
  if (!window.pywebview?.api?.delete_saved_email) return;
  try {
    await window.pywebview.api.delete_saved_email(path);
  } catch (e) {
    console.warn("Failed to delete managed saved email:", e);
  }
}

async function openDeliverableEmailRef(emailRef) {
  const ref = normalizeEmailRef(emailRef);
  if (!ref) return false;
  const localPath = getEmailRefLocalPath(ref);
  if (localPath && window.pywebview?.api?.open_path) {
    try {
      const res = await window.pywebview.api.open_path(localPath);
      if (res && res.status === "error") {
        throw new Error(res.message || "Unable to open email file.");
      }
      return true;
    } catch (e) {
      toast("Could not open linked email file.");
      return false;
    }
  }
  openExternalUrl(ref.url || ref.raw);
  return true;
}

function buildEmailRefFromRawInput(rawInput, source = "url") {
  const raw = String(rawInput || "").trim();
  if (!raw) return null;
  const isUrlLike =
    /^https?:\/\//i.test(raw) ||
    /^outlook:/i.test(raw) ||
    /^mailto:/i.test(raw) ||
    /^file:\/\//i.test(raw);
  const isPathLike = /^[A-Za-z]:[\\/]/.test(raw) || /^\\\\/.test(raw);
  if (!isUrlLike && !isPathLike && !isEmailFilePath(raw)) {
    return null;
  }
  const normalized = normalizeEmailRef({
    raw,
    url: toFileURL(raw),
    label: basename(raw) || "Email",
    source,
    savedAt: new Date().toISOString(),
  });
  if (!normalized) return null;
  if (source === "saved-file") normalized.source = "saved-file";
  return normalized;
}

function extractEmailUrlFromUriList(uriList) {
  const lines = String(uriList || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line && !line.startsWith("#"));
  return lines[0] || "";
}

function extractEmailUrlFromHtml(html) {
  if (!html) return "";
  try {
    const doc = new DOMParser().parseFromString(html, "text/html");
    const anchor = doc.querySelector("a[href]");
    return anchor?.getAttribute("href")?.trim() || "";
  } catch {
    return "";
  }
}

function extractEmailUrlFromText(text) {
  const s = String(text || "").trim();
  if (!s) return "";
  const match = s.match(
    /(https?:\/\/[^\s<>"']+|outlook:[^\s<>"']+|mailto:[^\s<>"']+|file:\/\/\/[^\s<>"']+)/i
  );
  if (match) return match[1].trim();
  return "";
}

function getTransferDataText(dt, type) {
  try {
    return dt?.getData?.(type) || "";
  } catch {
    return "";
  }
}

function extractEmailUrlFromDropData(dt) {
  const uriList =
    getTransferDataText(dt, "text/uri-list") ||
    getTransferDataText(dt, "URL") ||
    getTransferDataText(dt, "UniformResourceLocator") ||
    getTransferDataText(dt, "UniformResourceLocatorW");
  const fromUriList = extractEmailUrlFromUriList(uriList);
  if (fromUriList) return fromUriList;

  const html = getTransferDataText(dt, "text/html");
  const fromHtml = extractEmailUrlFromHtml(html);
  if (fromHtml) return fromHtml;

  const text =
    getTransferDataText(dt, "text/plain") ||
    getTransferDataText(dt, "Text") ||
    getTransferDataText(dt, "text/x-moz-url");
  return extractEmailUrlFromText(text);
}

function isSupportedEmailFile(file) {
  if (!file) return false;
  const name = String(file.name || "").toLowerCase();
  return name.endsWith(".msg") || name.endsWith(".eml");
}

function buildDeliverableEmailContext(deliverable, project, scope = "projects-tab") {
  return {
    projectId: String(project?.id || val("f_id") || "").trim(),
    projectName: String(project?.name || val("f_name") || "").trim(),
    deliverableId: String(deliverable?.id || "").trim(),
    deliverableName: String(deliverable?.name || "").trim(),
    scope,
  };
}

async function saveDroppedEmailFile(file, context = {}) {
  if (!isSupportedEmailFile(file)) {
    throw new Error("Only .msg or .eml files are supported.");
  }
  if (!window.pywebview?.api?.save_dropped_email) {
    throw new Error("Email file saving API is unavailable.");
  }
  const upload = {
    name: file.name || "email.msg",
    dataUrl: await readFileAsDataUrl(file),
  };
  const response = await window.pywebview.api.save_dropped_email(upload, context);
  if (!response || response.status !== "success") {
    throw new Error(response?.message || "Failed to save dropped email.");
  }
  const normalized = normalizeEmailRef(response.emailRef);
  if (!normalized) throw new Error("Invalid email reference returned by backend.");
  normalized.source = "saved-file";
  return normalized;
}

async function resolveEmailRefFromDrop(event, context = {}) {
  const dt = event?.dataTransfer;
  if (!dt) return { emailRef: null, reason: "no-data-transfer" };

  const files = Array.from(dt.files || []);
  const emailFile = files.find(isSupportedEmailFile);
  if (emailFile) {
    const emailRef = await saveDroppedEmailFile(emailFile, context);
    return { emailRef, source: "file-drop" };
  }

  const urlCandidate = extractEmailUrlFromDropData(dt);
  if (urlCandidate) {
    const emailRef = buildEmailRefFromRawInput(urlCandidate, "url");
    if (emailRef) {
      return { emailRef, source: "url-drop" };
    }
  }

  return { emailRef: null, reason: "unsupported-drop-data" };
}

async function promptForEmailLinkRef() {
  const raw = prompt("Paste an Outlook/OWA email link.");
  if (raw == null) return null;
  const emailRef = buildEmailRefFromRawInput(raw, "url");
  if (!emailRef) {
    toast("That does not look like a valid email link.");
    return null;
  }
  return emailRef;
}

async function chooseEmailFileRefFromPicker() {
  if (!window.pywebview?.api?.select_files) {
    toast("File picker is unavailable.");
    return null;
  }
  try {
    const res = await window.pywebview.api.select_files({
      allow_multiple: false,
      file_types: [
        "Email Files (*.msg;*.eml)",
        "Outlook Message (*.msg)",
        "Email Message (*.eml)",
      ],
    });
    if (res?.status !== "success" || !res.paths?.length) return null;
    return buildEmailRefFromRawInput(res.paths[0], "file");
  } catch (e) {
    toast("Unable to open file picker.");
    return null;
  }
}

async function requestEmailRefFromFallbackActions() {
  const choosePaste = confirm(
    "No email is linked yet.\nPress OK to paste an email link.\nPress Cancel to choose a .msg/.eml file."
  );
  if (choosePaste) {
    const pasted = await promptForEmailLinkRef();
    if (pasted) return pasted;
    const chooseFileInstead = confirm("No valid link was provided. Choose a .msg/.eml file instead?");
    if (!chooseFileInstead) return null;
  }
  return chooseEmailFileRefFromPicker();
}

function showEmailLinkFallbackGuidance() {
  toast("Could not read Outlook drop data. Use the email icon to paste a link or choose a .msg/.eml file.");
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
  const projectsFilters = document.querySelector(
    "#projects-panel .projects-filter-controls"
  );
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
  if (projectsFilters)
    document.documentElement.style.setProperty(
      "--projects-filters-height",
      `${projectsFilters.offsetHeight}px`
    );
};

const debouncedStickyOffsets = debounce(updateStickyOffsets, 150);
window.addEventListener("resize", debouncedStickyOffsets);

const PROJECTS_BACK_TO_TOP_SCROLL_THRESHOLD = 240;

const isProjectsTabActive = () => {
  const panel = document.getElementById("projects-panel");
  return Boolean(panel && !panel.hidden && panel.classList.contains("active"));
};

function updateProjectsBackToTopVisibility() {
  const backToTopBtn = document.getElementById("projectsBackToTopBtn");
  if (!backToTopBtn) return;
  const scrollTop = window.scrollY || document.documentElement.scrollTop || 0;
  const shouldShow =
    isProjectsTabActive() && scrollTop > PROJECTS_BACK_TO_TOP_SCROLL_THRESHOLD;
  backToTopBtn.classList.toggle("is-visible", shouldShow);
  backToTopBtn.setAttribute("aria-hidden", shouldShow ? "false" : "true");
}

function initProjectsBackToTop() {
  const backToTopBtn = document.getElementById("projectsBackToTopBtn");
  if (!backToTopBtn) return;
  backToTopBtn.addEventListener("click", () => {
    window.scrollTo({ top: 0, behavior: "smooth" });
  });
  window.addEventListener("scroll", updateProjectsBackToTopVisibility, {
    passive: true,
  });
  window.addEventListener("resize", updateProjectsBackToTopVisibility);
  updateProjectsBackToTopVisibility();
}

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
let pendingCadLaunchContext = null;
let modalEmailSession = {
  active: false,
  created: new Map(),
  deleteOnSave: new Map(),
};

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
  workroomAutoSelectCadFiles: true,
};
let hideNonPrimary = true;
let activeNoteTab = null;
let latestAppUpdate = null;
let currentStatsTimespan = "1Y";
let currentStatsAggregation = "month";
let lightingScheduleProjectIndex = null;
let lightingScheduleProjectQuery = "";
let lightingTemplateQuery = "";
let lightingScheduleSyncStatusMessage = "";
let title24ProjectIndex = null;
let title24ProjectQuery = "";
let title24ScopeOptionsDataset = null;
let title24ScopeOptionsLoadError = "";
let title24ScopeOptionsLoadingPromise = null;
let aiNoMatchState = {
  rawAiData: null,
  aiProject: null,
  query: "",
  selectedProjectIndex: -1,
  candidateProjectIndices: [],
  mode: "choice",
};

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
    const sunIcon = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="5"></circle><line x1="12" y1="1" x2="12" y2="3"></line><line x1="12" y1="21" x2="12" y2="23"></line><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"></line><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"></line><line x1="1" y1="12" x2="3" y2="12"></line><line x1="21" y1="12" x2="23" y2="12"></line><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"></line><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"></line></svg>`;
    const moonIcon = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"></path></svg>`;
    toggle.innerHTML = isLight ? moonIcon : sunIcon;
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

// Debounced save for inline edits
let saveTimeout = null;
function debouncedSave() {
  if (saveTimeout) clearTimeout(saveTimeout);
  saveTimeout = setTimeout(() => {
    save();
    saveTimeout = null;
  }, 500);
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
    userSettings.workroomAutoSelectCadFiles =
      userSettings.workroomAutoSelectCadFiles !== false;
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

function syncWorkroomCadRoutingInputs() {
  const enabled = userSettings.workroomAutoSelectCadFiles !== false;
  setCheckboxValue(
    "settings_workroomAutoSelectCadFiles",
    enabled
  );
  setCheckboxValue("workroom_modal_autoSelectCadFiles", enabled);
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
  syncWorkroomCadRoutingInputs();

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

const NUMBERED_DELIVERABLE_NAME_RULES = [
  { regex: /\bRFI\s*#?\s*0*(\d+)\b/i, label: "RFI" },
  { regex: /\bASR\s*#?\s*0*(\d+)\b/i, label: "ASR" },
  { regex: /\bPCC\s*#?\s*0*(\d+)\b/i, label: "PCC" },
  { regex: /\bBulletin\s*#?\s*0*(\d+)\b/i, label: "Bulletin" },
  { regex: /\bRevision\s*#?\s*0*(\d+)\b/i, label: "Revision" },
];

const BASE_DELIVERABLE_NAME_RULES = [
  { regex: /\bRFI\b/i, label: "RFI" },
  { regex: /\bASR\b/i, label: "ASR" },
  { regex: /\bPCC\b/i, label: "PCC" },
  { regex: /\bBulletin\b/i, label: "Bulletin" },
];

const DELIVERABLE_NAME_RULES = [
  { regex: /\b(?:DD\s*[- ]?\s*50|50\s*%?\s*DD)\b/i, label: "DD50" },
  { regex: /\b(?:DD\s*[- ]?\s*60|60\s*%?\s*DD)\b/i, label: "DD60" },
  { regex: /\b(?:DD\s*[- ]?\s*90|90\s*%?\s*DD)\b/i, label: "DD90" },
  { regex: /\b(?:CD\s*[- ]?\s*50|50\s*%?\s*CDS?)\b/i, label: "CD50" },
  { regex: /\b(?:CD\s*[- ]?\s*60|60\s*%?\s*CDS?)\b/i, label: "CD60" },
  { regex: /\b(?:CD\s*[- ]?\s*90|90\s*%?\s*CDS?)\b/i, label: "CD90" },
  { regex: /\b(?:CD\s*[- ]?\s*100|100\s*%?\s*CDS?)\b/i, label: "CD100" },
  { regex: /\bCDF\b/i, label: "CDF" },
  { regex: /\bIFP\b/i, label: "IFP" },
  { regex: /\bIFC\b/i, label: "IFC" },
  { regex: /\blighting\s+submittals?\b/i, label: "Lighting Submittal" },
  { regex: /\bcontrols?\s+submittals?\b/i, label: "Controls Submittal" },
  { regex: /\brecord\s+drawings?\b/i, label: "Record Drawings" },
  { regex: /\brecord\s+sets?\b/i, label: "Record Set" },
  { regex: /\bsite\s+survey\b/i, label: "Site Survey" },
  { regex: /\bsurvey\s+reports?\b/i, label: "Survey Report" },
  { regex: /\bsubmittals?\b/i, label: "Submittal" },
  { regex: /\bcoordination\b/i, label: "Coordination" },
  { regex: /\bmeeting\b/i, label: "Meeting" },
  { regex: /\brevision\b/i, label: "Revision" },
];

function normalizeDeliverableSequenceNumber(value) {
  const parsed = Number.parseInt(String(value || "").trim(), 10);
  if (!Number.isFinite(parsed)) return "";
  return String(parsed);
}

function extractDeliverableName(text) {
  if (!text) return "";
  const raw = String(text);

  for (const rule of NUMBERED_DELIVERABLE_NAME_RULES) {
    const match = raw.match(rule.regex);
    if (!match) continue;
    const sequence = normalizeDeliverableSequenceNumber(match[1]);
    if (sequence) return `${rule.label} #${sequence}`;
  }

  for (const rule of BASE_DELIVERABLE_NAME_RULES) {
    if (rule.regex.test(raw)) return rule.label;
  }

  for (const rule of DELIVERABLE_NAME_RULES) {
    if (rule.regex.test(raw)) return rule.label;
  }
  return "";
}

function guessDeliverableName(legacy) {
  const candidates = [];
  if (legacy?.deliverable) candidates.push(legacy.deliverable);
  if (Array.isArray(legacy?.tasks)) {
    legacy.tasks.forEach((t) => candidates.push(t?.text || t));
  }
  if (legacy?.notes) candidates.push(legacy.notes);
  if (legacy?.path) candidates.push(basename(legacy.path));
  if (legacy?.name) candidates.push(legacy.name);
  if (legacy?.nick) candidates.push(legacy.nick);
  for (const text of candidates) {
    const name = extractDeliverableName(text);
    if (name) return name;
  }
  return "Deliverable";
}

function normalizeWorkroomCadDiscipline(value, fallback = "") {
  const raw = String(value || "").trim();
  const caseInsensitiveMatch = WORKROOM_CAD_DISCIPLINES.find(
    (discipline) => discipline.toLowerCase() === raw.toLowerCase()
  );
  if (caseInsensitiveMatch) return caseInsensitiveMatch;
  const fallbackMatch = WORKROOM_CAD_DISCIPLINES.find(
    (discipline) => discipline.toLowerCase() === String(fallback || "").toLowerCase()
  );
  return fallbackMatch || "";
}

function normalizeDeliverable(deliverable = {}) {
  const emailRefs = normalizeEmailRefs(deliverable.emailRefs, deliverable.emailRef);
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
    emailRefs,
    emailRef: emailRefs[0] || null,
    workroomCadDiscipline: normalizeWorkroomCadDiscipline(
      deliverable.workroomCadDiscipline,
      ""
    ),
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
  const seedEmailRefs = normalizeEmailRefs(seed.emailRefs, seed.emailRef);
  return normalizeDeliverable({
    id: seed.id || createId("dlv"),
    name: seed.name || "",
    due: seed.due || "",
    notes: seed.notes || "",
    tasks: seed.tasks || [],
    statuses: seed.statuses || [],
    statusTags: seed.statusTags || [],
    status: seed.status || "",
    emailRefs: seedEmailRefs,
    emailRef: seedEmailRefs[0] || null,
    workroomCadDiscipline: seed.workroomCadDiscipline || "",
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
  if (!base.title24 && incoming.title24) base.title24 = incoming.title24;
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
    title24: legacy?.title24 || null,
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
    lightingSchedule: normalizeLightingSchedule(
      project.lightingSchedule || createDefaultLightingSchedule()
    ),
    title24: normalizeTitle24(project.title24 || createDefaultTitle24()),
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
      if (needsLightingScheduleMigration(item?.lightingSchedule)) {
        changed = true;
      }
      if (needsTitle24Migration(item?.title24)) {
        changed = true;
      }
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

async function pinUrgentDeliverables() {
  const confirmed = confirm(
    "Pin all projects with overdue incomplete deliverables or deliverables due today?"
  );
  if (!confirmed) return;

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let projectsPinned = 0;
  let deliverablesMatched = 0;

  db.forEach((project) => {
    const deliverables = getProjectDeliverables(project);
    const urgentIncompleteCount = deliverables.filter((deliverable) => {
      if (isFinished(deliverable)) return false;
      const due = parseDueStr(deliverable?.due);
      if (!due) return false;
      return isSameDay(due, today) || due < today;
    }).length;

    if (!urgentIncompleteCount) return;
    deliverablesMatched += urgentIncompleteCount;
    if (!project.pinned) {
      project.pinned = true;
      projectsPinned += 1;
    }
  });

  if (!deliverablesMatched) {
    toast("No urgent incomplete deliverables found.");
    return;
  }

  await save();
  render();
  toast(
    projectsPinned
      ? `Pinned ${projectsPinned} project${projectsPinned === 1 ? "" : "s"} with ${deliverablesMatched} urgent deliverable${deliverablesMatched === 1 ? "" : "s"}.`
      : `${deliverablesMatched} urgent deliverable${deliverablesMatched === 1 ? "" : "s"} already pinned.`
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

function setEmailButtonBusy(button, busy) {
  if (!button) return;
  button.classList.toggle("is-busy", !!busy);
  button.setAttribute("aria-busy", String(!!busy));
}

function ensureEmailButtonIcon(button) {
  if (!button || button.querySelector("svg")) return;
  button.textContent = "";
  button.appendChild(createIcon(MAIL_ICON_PATH, 14));
}

function updateDeliverableEmailButtonState(button, emailRef, slotIndex = 0) {
  if (!button) return;
  ensureEmailButtonIcon(button);
  const normalized = normalizeEmailRef(emailRef);
  const linked = !!normalized;
  const slotNumber = slotIndex + 1;
  button.classList.toggle("is-linked", linked);
  button.classList.toggle("is-empty", !linked);
  button.classList.remove("is-dragover");
  if (linked) {
    button.title = `Open linked email ${slotNumber} (Shift+Click to replace)`;
    button.setAttribute("aria-label", `Open linked email ${slotNumber}`);
  } else {
    button.title = `Link email ${slotNumber}`;
    button.setAttribute("aria-label", `Link email ${slotNumber}`);
  }
}

function getManagedEmailPathMap(emailRefs) {
  const map = new Map();
  normalizeEmailRefs(emailRefs).forEach((emailRef) => {
    if (!isManagedSavedEmailRef(emailRef)) return;
    const path = getEmailRefLocalPath(emailRef);
    if (!path) return;
    map.set(path.toLowerCase(), path);
  });
  return map;
}

function getAnyEmailPathMap(emailRefs) {
  const map = new Map();
  normalizeEmailRefs(emailRefs).forEach((emailRef) => {
    const path = getEmailRefLocalPath(emailRef);
    if (!path) return;
    map.set(path.toLowerCase(), path);
  });
  return map;
}

async function reconcileManagedEmailRefTransitions(previousRefs, nextRefs, options = {}) {
  const { modalCard = null } = options;
  const previousPaths = getManagedEmailPathMap(previousRefs);
  const nextManagedPaths = getManagedEmailPathMap(nextRefs);
  const nextAnyPaths = getAnyEmailPathMap(nextRefs);
  const removed = [];
  const added = [];

  previousPaths.forEach((path, key) => {
    if (!nextAnyPaths.has(key)) removed.push(path);
  });
  nextManagedPaths.forEach((path, key) => {
    if (!previousPaths.has(key)) added.push(path);
  });

  if (modalCard) {
    nextAnyPaths.forEach((_, key) => {
      modalEmailSession.deleteOnSave.delete(key);
    });
    for (const path of removed) {
      const key = path.toLowerCase();
      if (modalEmailSession.created.has(key)) {
        modalEmailSession.created.delete(key);
        await deleteManagedSavedEmailRef({ raw: path, source: "saved-file" });
      } else {
        modalEmailSession.deleteOnSave.set(key, path);
      }
    }
    for (const path of added) {
      const key = path.toLowerCase();
      modalEmailSession.created.set(key, path);
      modalEmailSession.deleteOnSave.delete(key);
    }
    return;
  }

  for (const path of removed) {
    await deleteManagedSavedEmailRef({ raw: path, source: "saved-file" });
  }
}

function beginModalEmailSession() {
  modalEmailSession.active = true;
  modalEmailSession.created = new Map();
  modalEmailSession.deleteOnSave = new Map();
}

async function flushModalEmailSession(committed) {
  if (!modalEmailSession.active) return;
  const created = [...modalEmailSession.created.values()];
  const deleteOnSave = [...modalEmailSession.deleteOnSave.values()];
  modalEmailSession.active = false;
  modalEmailSession.created = new Map();
  modalEmailSession.deleteOnSave = new Map();

  if (committed) {
    for (const path of deleteOnSave) {
      await deleteManagedSavedEmailRef({ raw: path, source: "saved-file" });
    }
    return;
  }

  for (const path of created) {
    await deleteManagedSavedEmailRef({ raw: path, source: "saved-file" });
  }
}

async function stageModalManagedEmailRefForRemoval(emailRef) {
  if (!isManagedSavedEmailRef(emailRef)) return;
  const path = getEmailRefLocalPath(emailRef);
  if (!path) return;
  if (!modalEmailSession.active) {
    await deleteManagedSavedEmailRef(emailRef);
    return;
  }
  const key = path.toLowerCase();
  if (modalEmailSession.created.has(key)) {
    modalEmailSession.created.delete(key);
    await deleteManagedSavedEmailRef(emailRef);
  } else {
    modalEmailSession.deleteOnSave.set(key, path);
  }
}

async function stageModalManagedEmailRefsForRemoval(emailRefs) {
  const refs = normalizeEmailRefs(emailRefs);
  for (const emailRef of refs) {
    await stageModalManagedEmailRefForRemoval(emailRef);
  }
}

async function applyDeliverableEmailRefAtIndex(deliverable, slotIndex, nextRef, options = {}) {
  const {
    persistNow = false,
    modalCard = null,
  } = options;
  const normalizedNext = normalizeEmailRef(nextRef);
  if (!deliverable || !normalizedNext) return false;
  if (!Number.isInteger(slotIndex)) return false;
  if (slotIndex < 0 || slotIndex >= MAX_DELIVERABLE_EMAIL_REFS) return false;

  const currentRefs = syncDeliverableEmailRefs(deliverable).slice();
  const currentLength = currentRefs.length;
  if (slotIndex > currentLength) return false;
  if (slotIndex === currentLength && currentLength >= MAX_DELIVERABLE_EMAIL_REFS) return false;
  if (slotIndex < currentLength && sameEmailRef(currentRefs[slotIndex], normalizedNext)) return false;

  const nextRefs = currentRefs.slice();
  if (slotIndex === currentLength) {
    nextRefs.push(normalizedNext);
  } else {
    nextRefs[slotIndex] = normalizedNext;
  }

  const normalizedNextRefs = normalizeEmailRefs(nextRefs);
  if (sameEmailRefList(currentRefs, normalizedNextRefs)) return false;
  await reconcileManagedEmailRefTransitions(currentRefs, normalizedNextRefs, { modalCard });

  deliverable.emailRefs = normalizedNextRefs;
  deliverable.emailRef = normalizedNextRefs[0] || null;
  if (modalCard) setDeliverableCardEmailRefs(modalCard, normalizedNextRefs);
  if (persistNow) await save();
  return true;
}

async function removeDeliverableEmailRefAtIndex(deliverable, slotIndex, options = {}) {
  const {
    persistNow = false,
    modalCard = null,
  } = options;
  if (!deliverable || !Number.isInteger(slotIndex)) return false;
  if (slotIndex < 0 || slotIndex >= MAX_DELIVERABLE_EMAIL_REFS) return false;
  const currentRefs = syncDeliverableEmailRefs(deliverable).slice();
  if (slotIndex >= currentRefs.length) return false;

  const nextRefs = currentRefs.filter((_, index) => index !== slotIndex);
  if (sameEmailRefList(currentRefs, nextRefs)) return false;
  await reconcileManagedEmailRefTransitions(currentRefs, nextRefs, { modalCard });

  deliverable.emailRefs = nextRefs;
  deliverable.emailRef = nextRefs[0] || null;
  if (modalCard) setDeliverableCardEmailRefs(modalCard, nextRefs);
  if (persistNow) await save();
  return true;
}

function renderDeliverableEmailSlots(container, deliverable, options = {}) {
  if (!container || !deliverable) return;
  const {
    project = null,
    persistNow = false,
    modalCard = null,
    scope = "projects-tab",
  } = options;

  syncDeliverableEmailRefs(deliverable);
  container.classList.add("deliverable-email-slots");
  container.innerHTML = "";
  const emailRefs = deliverable.emailRefs || [];
  const slotCount = emailRefs.length >= MAX_DELIVERABLE_EMAIL_REFS
    ? MAX_DELIVERABLE_EMAIL_REFS
    : Math.max(1, emailRefs.length + 1);

  const rerender = () =>
    renderDeliverableEmailSlots(container, deliverable, {
      project,
      persistNow,
      modalCard,
      scope,
    });

  for (let slotIndex = 0; slotIndex < slotCount; slotIndex++) {
    const currentRef = emailRefs[slotIndex] || null;
    const slot = el("div", { className: "deliverable-email-slot" });
    const button = el("button", {
      className: "deliverable-email-btn",
      type: "button",
    });

    updateDeliverableEmailButtonState(button, currentRef, slotIndex);

    const applyRef = async (nextRef) => {
      const changed = await applyDeliverableEmailRefAtIndex(
        deliverable,
        slotIndex,
        nextRef,
        {
          persistNow,
          modalCard,
        }
      );
      if (changed) rerender();
      return changed;
    };

    button.onclick = async (e) => {
      e.stopPropagation();
      if (button.classList.contains("is-busy")) return;
      const refs = syncDeliverableEmailRefs(deliverable);
      const slotRef = refs[slotIndex] || null;
      if (slotRef && !e.shiftKey) {
        await openDeliverableEmailRef(slotRef);
        return;
      }
      const nextRef = await requestEmailRefFromFallbackActions();
      if (!nextRef) return;
      await applyRef(nextRef);
    };

    button.addEventListener("dragover", (e) => {
      e.preventDefault();
      e.stopPropagation();
      if (button.classList.contains("is-busy")) return;
      button.classList.add("is-dragover");
      if (e.dataTransfer) e.dataTransfer.dropEffect = "copy";
    });

    button.addEventListener("dragleave", (e) => {
      e.preventDefault();
      e.stopPropagation();
      button.classList.remove("is-dragover");
    });

    button.addEventListener("drop", async (e) => {
      e.preventDefault();
      e.stopPropagation();
      button.classList.remove("is-dragover");
      if (button.classList.contains("is-busy")) return;
      setEmailButtonBusy(button, true);
      try {
        const context = buildDeliverableEmailContext(deliverable, project, scope);
        const resolved = await resolveEmailRefFromDrop(e, context);
        if (!resolved.emailRef) {
          showEmailLinkFallbackGuidance();
          return;
        }
        await applyRef(resolved.emailRef);
      } catch (error) {
        console.warn("Failed to resolve email from drop:", error);
        toast(error?.message || "Unable to process dropped email.");
        showEmailLinkFallbackGuidance();
      } finally {
        setEmailButtonBusy(button, false);
        if (button.isConnected) {
          const refs = syncDeliverableEmailRefs(deliverable);
          updateDeliverableEmailButtonState(button, refs[slotIndex] || null, slotIndex);
        }
      }
    });

    slot.appendChild(button);

    if (currentRef) {
      const removeBtn = el("button", {
        className: "deliverable-email-remove",
        type: "button",
        title: `Remove linked email ${slotIndex + 1}`,
        "aria-label": `Remove linked email ${slotIndex + 1}`,
        textContent: "x",
      });
      removeBtn.onclick = async (e) => {
        e.stopPropagation();
        if (button.classList.contains("is-busy")) return;
        const changed = await removeDeliverableEmailRefAtIndex(
          deliverable,
          slotIndex,
          {
            persistNow,
            modalCard,
          }
        );
        if (changed) rerender();
      };
      slot.appendChild(removeBtn);
    }

    container.appendChild(slot);
  }
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
    const badgeText = humanDate(deliverable.due).replace(/\//g, "/").toUpperCase();
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
        ? `Overdue – due ${humanDate(deliverable.due)}. Click to change due date.`
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

  // Right section: email + expand/contract controls
  const actions = el("div", { className: "deliverable-header-actions" });
  const emailSlots = el("div", {
    className: "deliverable-email-slots",
    "aria-label": "Deliverable email links",
  });
  renderDeliverableEmailSlots(emailSlots, deliverable, {
    project,
    persistNow: true,
    modalCard: null,
    scope: "projects-tab",
  });

  const expandToggle = createExpandToggle(card);
  actions.append(emailSlots, expandToggle);
  header.appendChild(actions);

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

function buildAiDeliverableFromData(rawAiData = {}) {
  return createDeliverable({
    name: guessDeliverableName(rawAiData),
    due: rawAiData?.due || "",
    notes: rawAiData?.notes || "",
    tasks: rawAiData?.tasks || [],
  });
}

function openAiCreateNewProject(rawAiData = {}) {
  openNew();
  fillForm(rawAiData || {});
}

function resetAiNoMatchState() {
  aiNoMatchState = {
    rawAiData: null,
    aiProject: null,
    query: "",
    selectedProjectIndex: -1,
    candidateProjectIndices: [],
    mode: "choice",
  };
}

function setAiNoMatchDialogMode(mode = "choice") {
  aiNoMatchState.mode = mode === "search" ? "search" : "choice";
  const choiceSection = document.getElementById("aiNoMatchChoiceSection");
  const searchSection = document.getElementById("aiNoMatchSearchSection");
  if (choiceSection) choiceSection.hidden = aiNoMatchState.mode !== "choice";
  if (searchSection) searchSection.hidden = aiNoMatchState.mode !== "search";
}

function getAiNoMatchCandidates(query = "") {
  const normalizedQuery = normalizeProjectMatchValue(query);
  return db
    .map((project, index) => ({ index, project: normalizeProject(project) }))
    .filter(({ project }) => !!project)
    .filter(({ project }) => {
      if (!normalizedQuery) return true;
      const name = normalizeProjectMatchValue(project?.name);
      const nick = normalizeProjectMatchValue(project?.nick);
      return (
        (name && name.includes(normalizedQuery)) ||
        (nick && nick.includes(normalizedQuery))
      );
    });
}

function renderAiNoMatchProjectOptions() {
  const list = document.getElementById("aiNoMatchProjectList");
  const addBtn = document.getElementById("aiNoMatchAddBtn");
  if (!list) return;

  const candidates = getAiNoMatchCandidates(aiNoMatchState.query);
  aiNoMatchState.candidateProjectIndices = candidates.map(
    ({ index }) => index
  );
  if (
    aiNoMatchState.selectedProjectIndex >= 0 &&
    !aiNoMatchState.candidateProjectIndices.includes(
      aiNoMatchState.selectedProjectIndex
    )
  ) {
    aiNoMatchState.selectedProjectIndex = -1;
  }

  list.innerHTML = "";
  candidates.forEach(({ index, project }) => {
    const option = el("button", {
      className:
        "ts-project-option ai-no-match-project-option" +
        (aiNoMatchState.selectedProjectIndex === index ? " is-selected" : ""),
      type: "button",
    });
    option.append(
      el("div", {
        className: "ai-no-match-project-option-name",
        textContent:
          project.name ||
          project.nick ||
          project.id ||
          `Project ${index + 1}`,
      }),
      el("div", {
        className: "ai-no-match-project-option-meta",
        textContent:
          `Nickname: ${project.nick || "--"} | ID: ${project.id || "--"} | ` +
          `Deliverables: ${getProjectDeliverables(project).length}`,
      })
    );
    option.onclick = () => {
      aiNoMatchState.selectedProjectIndex = index;
      renderAiNoMatchProjectOptions();
    };
    list.appendChild(option);
  });

  if (!candidates.length) {
    const message = db.length
      ? "No matching projects found."
      : "No projects exist yet. Create a new project from AI details.";
    list.innerHTML = `<p class="muted" style="padding: 1rem; text-align: center;">${message}</p>`;
  }

  if (addBtn) addBtn.disabled = aiNoMatchState.selectedProjectIndex < 0;
}

function getAiNoMatchDialogState() {
  const dlg = document.getElementById("aiNoMatchDlg");
  const addBtn = document.getElementById("aiNoMatchAddBtn");
  return {
    open: !!dlg?.open,
    mode: aiNoMatchState.mode,
    query: aiNoMatchState.query,
    selectedProjectIndex: aiNoMatchState.selectedProjectIndex,
    candidateProjectIndices: [...aiNoMatchState.candidateProjectIndices],
    resultCount: aiNoMatchState.candidateProjectIndices.length,
    canAdd: !!addBtn && !addBtn.disabled,
  };
}

function openAiNoMatchResolution(rawAiData = {}, aiProject = null) {
  const dlg = document.getElementById("aiNoMatchDlg");
  if (!dlg) return { open: false };

  aiNoMatchState.rawAiData = rawAiData || {};
  aiNoMatchState.aiProject = aiProject || normalizeProject(rawAiData || {});
  aiNoMatchState.query = String(
    aiNoMatchState.aiProject?.name || aiNoMatchState.aiProject?.nick || ""
  ).trim();
  aiNoMatchState.selectedProjectIndex = -1;
  aiNoMatchState.candidateProjectIndices = [];
  setAiNoMatchDialogMode("choice");

  const searchInput = document.getElementById("aiNoMatchSearchInput");
  if (searchInput) searchInput.value = aiNoMatchState.query;
  renderAiNoMatchProjectOptions();
  showDialog(dlg);
  return getAiNoMatchDialogState();
}

function openAiNoMatchSearch() {
  const searchInput = document.getElementById("aiNoMatchSearchInput");
  setAiNoMatchDialogMode("search");
  if (searchInput) {
    searchInput.value = aiNoMatchState.query;
    searchInput.focus();
    searchInput.select();
  }
  renderAiNoMatchProjectOptions();
  return getAiNoMatchDialogState();
}

function setAiNoMatchSearchQuery(query = "") {
  aiNoMatchState.query = String(query || "");
  const searchInput = document.getElementById("aiNoMatchSearchInput");
  if (searchInput && searchInput.value !== aiNoMatchState.query) {
    searchInput.value = aiNoMatchState.query;
  }
  renderAiNoMatchProjectOptions();
  return getAiNoMatchDialogState();
}

function selectAiNoMatchProject(indexOrPosition) {
  const value = Number(indexOrPosition);
  if (!Number.isInteger(value)) return false;

  let resolvedIndex = -1;
  if (aiNoMatchState.candidateProjectIndices.includes(value)) {
    resolvedIndex = value;
  } else if (
    value >= 0 &&
    value < aiNoMatchState.candidateProjectIndices.length
  ) {
    resolvedIndex = aiNoMatchState.candidateProjectIndices[value];
  }
  if (resolvedIndex < 0 || !db[resolvedIndex]) return false;

  aiNoMatchState.selectedProjectIndex = resolvedIndex;
  renderAiNoMatchProjectOptions();
  return true;
}

function closeAiNoMatchDialog() {
  const dlg = document.getElementById("aiNoMatchDlg");
  if (dlg?.open) dlg.close();
  resetAiNoMatchState();
}

function addAiDeliverableToProject(projectIndex, aiProject, rawAiData) {
  if (!Number.isInteger(projectIndex) || projectIndex < 0 || !db[projectIndex]) {
    return false;
  }

  const snapshot = JSON.parse(JSON.stringify(db[projectIndex]));
  const target = normalizeProject(db[projectIndex]);
  if (!target) return false;

  const newDeliverable = buildAiDeliverableFromData(rawAiData);
  if (!target.path && aiProject?.path) target.path = aiProject.path;
  if (!target.id && aiProject?.id) target.id = aiProject.id;
  target.deliverables.push(newDeliverable);
  target.overviewDeliverableId = newDeliverable.id;
  db[projectIndex] = target;
  editIndex = projectIndex;
  _aiMatchSnapshot = { index: projectIndex, data: snapshot };
  document.getElementById("dlgTitle").textContent =
    `Edit Project \u2014 ${target.id || "Untitled"}`;
  document.getElementById("btnSaveProject").textContent = "Save Changes";
  fillForm(target);
  document.getElementById("editDlg").showModal();
  return true;
}

function applyAiToMatchedProject(match, aiProject, rawAiData) {
  if (!match || !Number.isInteger(match.index)) return false;
  const applied = addAiDeliverableToProject(match.index, aiProject, rawAiData);
  if (!applied) return false;
  const confidenceText = Number.isFinite(match.score)
    ? `${Math.round(match.score * 100)}% confidence`
    : "match confirmed";
  toast(`Matched existing project (${match.method || "auto"}, ${confidenceText}).`);
  return true;
}

function confirmAiNoMatchAddToProject() {
  const selectedIndex = aiNoMatchState.selectedProjectIndex;
  if (!Number.isInteger(selectedIndex) || selectedIndex < 0 || !db[selectedIndex]) {
    return false;
  }
  const aiProject = aiNoMatchState.aiProject || normalizeProject(aiNoMatchState.rawAiData || {});
  const rawAiData = aiNoMatchState.rawAiData || {};
  closeAiNoMatchDialog();
  const applied = addAiDeliverableToProject(selectedIndex, aiProject, rawAiData);
  if (applied) toast("Added deliverable to selected project.");
  return applied;
}

function handleAiProjectResult(rawAiData) {
  const aiProject = normalizeProject(rawAiData || {});
  if (!aiProject) {
    openAiCreateNewProject(rawAiData || {});
    return { branch: "new-project" };
  }
  const match = findBestProjectMatch(aiProject);
  if (match) {
    applyAiToMatchedProject(match, aiProject, rawAiData || {});
    return { branch: "matched", match };
  }
  const dialogState = openAiNoMatchResolution(rawAiData || {}, aiProject);
  if (!dialogState.open) {
    openAiCreateNewProject(rawAiData || {});
    return { branch: "new-project" };
  }
  return { branch: "no-match" };
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

function createContextSnippet(text, q, contextChars = 60) {
  const span = el("span", { className: "search-context-snippet" });
  const lower = text.toLowerCase();
  const idx = lower.indexOf(q);
  if (idx === -1) {
    span.textContent = text.slice(0, contextChars * 2) + (text.length > contextChars * 2 ? "..." : "");
    return span;
  }
  const start = Math.max(0, idx - contextChars);
  const end = Math.min(text.length, idx + q.length + contextChars);
  const snippet = (start > 0 ? "..." : "") + text.slice(start, end) + (end < text.length ? "..." : "");
  span.appendChild(highlightText(snippet, q));
  return span;
}

function buildMatchContextRow(q, project, context) {
  const tr = el("tr", { className: "search-context-row" });
  const td = el("td", { colSpan: 7 });
  const container = el("div", { className: "search-context" });

  if (context.projectFields.includes("notes") && project.notes) {
    const item = el("div", { className: "search-context-item" });
    item.append(
      el("span", { className: "search-context-label", textContent: "Project Notes: " }),
      createContextSnippet(project.notes, q)
    );
    container.appendChild(item);
  }

  context.deliverables.forEach((dCtx) => {
    const d = dCtx.deliverable;

    if (dCtx.fields.includes("name")) {
      const item = el("div", { className: "search-context-item" });
      const label = el("span", { className: "search-context-label", textContent: "Deliverable: " });
      item.append(label);
      item.appendChild(highlightText(d.name, q));
      container.appendChild(item);
    }

    if (dCtx.fields.includes("notes") && d.notes) {
      const item = el("div", { className: "search-context-item" });
      item.append(
        el("span", { className: "search-context-label", textContent: d.name + " Notes: " }),
        createContextSnippet(d.notes, q)
      );
      container.appendChild(item);
    }

    dCtx.matchingTasks.forEach((taskText) => {
      const item = el("div", { className: "search-context-item" });
      item.append(
        el("span", { className: "search-context-label", textContent: d.name + " Task: " })
      );
      item.appendChild(highlightText(taskText, q));
      container.appendChild(item);
    });
  });

  if (!container.children.length) return null;
  td.appendChild(container);
  tr.appendChild(td);
  return tr;
}

function render() {
  const tbody = document.getElementById("tbody");
  const emptyState = document.getElementById("emptyState");
  tbody.innerHTML = "";

  const q = val("search").toLowerCase();
  const matchContextMap = new Map();

  let items = db.filter((p) => {
    if (q) {
      const ctx = getMatchContext(q, p);
      if (!ctx) return false;
      matchContextMap.set(p, ctx);
    }
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

  const appendSectionSeparator = (label) => {
    const sep = el("tr", { className: "week-separator-row" });
    sep.appendChild(el("td", { colSpan: 7 }, [
      el("div", { className: "week-separator" }, [
        el("span", { className: "week-separator-label", textContent: label })
      ])
    ]));
    tbody.appendChild(sep);
  };

  let lastWeekKey = null;
  let pinnedSectionShown = false;
  items.forEach((p) => {
    const isPinnedProject = !!p?.pinned;
    const projectDue = getProjectSortKey(p);
    const weekKey = projectDue ? formatWeekKey(projectDue) : "no-date";
    if (isPinnedProject) {
      if (!pinnedSectionShown) {
        appendSectionSeparator("Pinned Projects");
        pinnedSectionShown = true;
      }
    } else {
      if (lastWeekKey === null || weekKey !== lastWeekKey) {
        const weekLabel = projectDue ? formatWeekDisplay(projectDue) : "";
        appendSectionSeparator(weekLabel);
      }
      lastWeekKey = weekKey;
    }

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
        title: "Project Workroom",
        ariaLabel: "Open project workroom",
        path: WORKROOM_ICON_PATH,
        onClick: () => openChecklistModal(idx),
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

    if (q && matchContextMap.has(p)) {
      const contextRow = buildMatchContextRow(q, p, matchContextMap.get(p));
      if (contextRow) tbody.appendChild(contextRow);
    }
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

function getMatchContext(q, p) {
  if (!q) return null;
  const str = (val) => (val || "").toLowerCase();
  const context = { projectFields: [], deliverables: [] };
  let hasMatch = false;

  if (str(p.id).includes(q)) { context.projectFields.push("id"); hasMatch = true; }
  if (str(p.name).includes(q)) { context.projectFields.push("name"); hasMatch = true; }
  if (str(p.nick).includes(q)) { context.projectFields.push("nick"); hasMatch = true; }
  if (str(p.notes).includes(q)) { context.projectFields.push("notes"); hasMatch = true; }

  getProjectDeliverables(p).forEach((d) => {
    const dCtx = { deliverable: d, fields: [], matchingTasks: [] };
    if (str(d.name).includes(q)) dCtx.fields.push("name");
    if (str(d.notes).includes(q)) dCtx.fields.push("notes");
    if (str(d.due).includes(q)) dCtx.fields.push("due");
    (d.tasks || []).forEach((t) => {
      if (str(t.text).includes(q)) dCtx.matchingTasks.push(t.text);
    });
    const hasStatusMatch = (d.statuses || []).some((s) => str(s).includes(q));
    if (dCtx.fields.length || dCtx.matchingTasks.length || hasStatusMatch) {
      context.deliverables.push(dCtx);
      hasMatch = true;
    }
  });

  return hasMatch ? context : null;
}

function matches(q, p) {
  if (!q) return true;
  return getMatchContext(q, p) !== null;
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
    lightingSchedule: createDefaultLightingSchedule(),
    title24: createDefaultTitle24(),
  };
}

function openEdit(i) {
  editIndex = i;
  beginModalEmailSession();
  const p = db[i];
  document.getElementById("dlgTitle").textContent = `Edit Project — ${p.id || "Untitled"
    }`;
  document.getElementById("btnSaveProject").textContent = "Save Changes";
  fillForm(p);
  document.getElementById("editDlg").showModal();
}
function openNew() {
  editIndex = -1;
  beginModalEmailSession();
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
  flushModalEmailSession(true);
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

function getDeliverableCardEmailRefs(card) {
  if (!card) return [];
  const raw = card.dataset.emailRefs || "";
  if (raw) {
    try {
      return normalizeEmailRefs(JSON.parse(raw));
    } catch {
      /* fall through to legacy field */
    }
  }
  const legacyRaw = card.dataset.emailRef || "";
  if (!legacyRaw) return [];
  try {
    return normalizeEmailRefs(null, JSON.parse(legacyRaw));
  } catch {
    return normalizeEmailRefs(null, legacyRaw);
  }
}

function setDeliverableCardEmailRefs(card, emailRefs) {
  if (!card) return;
  const normalized = normalizeEmailRefs(emailRefs);
  if (!normalized.length) {
    delete card.dataset.emailRefs;
    delete card.dataset.emailRef;
    return;
  }
  card.dataset.emailRefs = JSON.stringify(normalized);
  card.dataset.emailRef = JSON.stringify(normalized[0]);
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
  syncDeliverableEmailRefs(deliverable);
  setDeliverableCardEmailRefs(card, deliverable.emailRefs);

  card.querySelector(".d-name").value = deliverable.name || "";
  card.querySelector(".d-due").value = deliverable.due || "";
  card.querySelector(".d-notes").value = deliverable.notes || "";

  const primaryInput = card.querySelector(".d-primary");
  primaryInput.name = "primaryDeliverable";
  if (primaryId && deliverableId === primaryId) {
    primaryInput.checked = true;
  }

  const emailSlots = card.querySelector(".deliverable-email-slots");
  if (emailSlots) {
    renderDeliverableEmailSlots(emailSlots, deliverable, {
      project: null,
      persistNow: false,
      modalCard: card,
      scope: "edit-modal",
    });
  }

  // No secondary action needed on primary change anymore

  const taskList = card.querySelector(".deliverable-task-list");
  (deliverable.tasks || []).map(normalizeTask).forEach((task) => {
    addTaskRowFrom(taskList, task);
  });

  card.querySelector(".d-add-task").onclick = () =>
    addTaskRowFrom(taskList, {});
  card.querySelector(".btn-remove").onclick = async () => {
    const cardRefs = getDeliverableCardEmailRefs(card);
    await stageModalManagedEmailRefsForRemoval(
      cardRefs.length ? cardRefs : deliverable.emailRefs
    );
    card.remove();
  };

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
  const baseTitle24 =
    editIndex >= 0 && db[editIndex] ? db[editIndex].title24 : null;
  const existingProject = editIndex >= 0 && db[editIndex] ? db[editIndex] : null;
  const lightingSchedule = normalizeLightingSchedule(
    baseSchedule ?? createDefaultLightingSchedule()
  );
  const title24 = normalizeTitle24(baseTitle24 ?? createDefaultTitle24());
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
    title24,
  };

  document.querySelectorAll("#deliverableList .deliverable-card").forEach((card) => {
    const deliverableId = card.dataset.deliverableId || createId("dlv");
    const name = card.querySelector(".d-name").value.trim();
    const due = card.querySelector(".d-due").value.trim();
    const notes = card.querySelector(".d-notes").value;
    const statuses = readStatusPickerFrom(
      card.querySelector(".deliverable-status")
    );
    const emailRefs = getDeliverableCardEmailRefs(card);

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
      emailRefs,
      emailRef: emailRefs[0] || null,
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
      title24: createDefaultTitle24(),
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
      textContent: "x",
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

function focusAndSelectNoteSnippet(textarea, noteText, options = {}) {
  if (!textarea || !noteText) return;
  const { scrollWindow = false } = options;
  if (scrollWindow) {
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  const currentVal = textarea.value || "";
  const start = currentVal.indexOf(noteText);
  if (start === -1) {
    toast("Note content changed. Please refresh search.");
    return;
  }
  const end = start + noteText.length;

  textarea.blur();
  textarea.setSelectionRange(start, end);
  textarea.focus();

  const mirror = document.createElement("div");
  const style = window.getComputedStyle(textarea);

  mirror.style.visibility = "hidden";
  mirror.style.position = "absolute";
  mirror.style.top = "-9999px";
  mirror.style.whiteSpace = "pre-wrap";
  mirror.style.wordWrap = "break-word";

  ["fontFamily", "fontSize", "fontWeight", "lineHeight", "letterSpacing", "padding"].forEach(
    (prop) => {
      mirror.style[prop] = style[prop];
    }
  );

  mirror.style.boxSizing = "border-box";
  mirror.style.width = `${textarea.clientWidth}px`;
  mirror.style.border = "none";
  mirror.textContent = currentVal.substring(0, start);

  document.body.appendChild(mirror);
  const targetY = mirror.clientHeight;
  document.body.removeChild(mirror);

  textarea.scrollTop = Math.max(0, targetY - textarea.clientHeight * 0.3);
}

function createNoteSearchResultItem(noteText, onEdit) {
  const item = el("div", {
    className: "note-result-item",
    style: "position: relative;",
  });
  const contentDiv = el("div", { className: "note-result-content" });
  contentDiv.append(el("div", { className: "snippet", textContent: noteText }));

  const copyIcon = el("button", {
    className: "note-action-icon copy-icon",
    textContent: "\uD83D\uDCCB",
    title: "Copy",
    type: "button",
  });
  copyIcon.onclick = () => {
    navigator.clipboard.writeText(noteText).then(() => toast("Copied!"));
  };

  const editIcon = el("button", {
    className: "note-action-icon edit-icon",
    textContent: "\u270F\uFE0F",
    title: "Edit",
    type: "button",
  });
  editIcon.onclick = () => {
    if (typeof onEdit === "function") onEdit();
  };

  item.append(contentDiv, copyIcon, editIcon);
  return item;
}

function renderNoteSearchResults() {
  const query = val("notesSearch").toLowerCase();
  const resultsContainer = document.getElementById("notesSearchResults");
  if (!resultsContainer) return;
  resultsContainer.innerHTML = "";

  if (!query || !activeNoteTab) return;

  const queryWords = query.split(" ").filter((w) => w);
  if (!queryWords.length) return;

  const content = notesDb[activeNoteTab];
  if (!content) return;

  const notes = content.split(/\n\s*\n/).filter((note) => note.trim() !== "");

  notes.forEach((noteText) => {
    const lowerNoteText = noteText.toLowerCase();
    if (!queryWords.every((word) => lowerNoteText.includes(word))) return;

    resultsContainer.append(
      createNoteSearchResultItem(noteText, () => {
        const textarea = document.getElementById("notesTextarea");
        focusAndSelectNoteSnippet(textarea, noteText, { scrollWindow: true });
      })
    );
  });
}

function renderWorkroomNoteSearchResults() {
  const searchInput = document.getElementById("workroomNotesSearchInput");
  const resultsContainer = document.getElementById("workroomNotesSearchResults");
  const workroomTextarea = document.getElementById("workroomGeneralNotesTextarea");
  if (!searchInput || !resultsContainer || !workroomTextarea) return;

  const query = (searchInput.value || "").trim().toLowerCase();
  resultsContainer.innerHTML = "";
  if (!query || !activeNoteTab) return;

  const queryWords = query.split(" ").filter(Boolean);
  if (!queryWords.length) return;

  const content = notesDb[activeNoteTab];
  if (!content) return;

  const notes = content.split(/\n\s*\n/).filter((note) => note.trim() !== "");
  let matchCount = 0;
  notes.forEach((noteText) => {
    const lowerNoteText = noteText.toLowerCase();
    if (!queryWords.every((word) => lowerNoteText.includes(word))) return;
    matchCount += 1;

    resultsContainer.append(
      createNoteSearchResultItem(noteText, () => {
        focusAndSelectNoteSnippet(workroomTextarea, noteText);
      })
    );
  });

  if (matchCount === 0) {
    resultsContainer.innerHTML =
      "<p class='muted tiny' style='text-align:center;padding:0.75rem'>No results found.</p>";
  }
}
// ===================== CHECKLISTS RENDERING =====================

let activeChecklistTabId = null;
const debouncedSaveChecklists = debounce(saveChecklists, 500);

function renderChecklistTabs() {
  const container = document.getElementById("checklistsTabsContainer");
  container.innerHTML = "";

  checklistsDb.checklists.forEach((checklist) => {
    const btn = el("button", {
      className: `inner-tab-btn ${checklist.id === activeChecklistTabId ? "active" : ""}`,
      textContent: checklist.name,
      onclick: () => {
        activeChecklistTabId = checklist.id;
        renderChecklistTabs();
        renderChecklistItems();
      },
    });

    // Add remove icon for non-locked checklists
    if (!checklist.isLocked) {
      const delIcon = el("span", {
        className: "tab-delete-icon",
        textContent: "×",
        title: "Delete Checklist",
        onclick: (e) => {
          e.stopPropagation();
          if (confirm(`Permanently delete checklist "${checklist.name}"?`)) {
            deleteChecklist(checklist.id);
            if (activeChecklistTabId === checklist.id) {
              activeChecklistTabId = checklistsDb.checklists[0]?.id || null;
            }
            renderChecklistTabs();
            renderChecklistItems();
          }
        },
      });
      btn.appendChild(delIcon);
    }

    container.appendChild(btn);
  });

  // Add new checklist button
  const addBtn = el("button", {
    className: "add-tab-btn",
    textContent: "+",
    title: "Create New Checklist",
    onclick: () => {
      const name = prompt("Enter name for new checklist:");
      if (name && name.trim()) {
        const newChecklist = createChecklist(name.trim());
        activeChecklistTabId = newChecklist.id;
        renderChecklistTabs();
        renderChecklistItems();
      }
    },
  });
  container.appendChild(addBtn);

  renderChecklistItems();
}

function renderChecklistItems() {
  const container = document.getElementById("checklistItemsContainer");
  const nameInput = document.getElementById("checklistName");
  const resetBtn = document.getElementById("resetChecklistBtn");
  const deleteBtn = document.getElementById("deleteChecklistBtn");
  const addItemBtn = document.getElementById("addChecklistItemBtn");

  container.innerHTML = "";

  const checklist = getChecklistById(activeChecklistTabId);

  if (!checklist) {
    nameInput.value = "";
    nameInput.disabled = true;
    resetBtn.style.display = "none";
    deleteBtn.style.display = "none";
    addItemBtn.disabled = true;
    container.innerHTML = '<p class="muted">Select or create a checklist to get started.</p>';
    return;
  }

  nameInput.value = checklist.name;
  nameInput.disabled = checklist.isLocked;
  resetBtn.style.display = checklist.isDefault ? "inline-flex" : "none";
  deleteBtn.style.display = checklist.isLocked ? "none" : "inline-flex";
  addItemBtn.disabled = false;

  // Render items
  checklist.items.forEach((item, index) => {
    const row = el("div", { className: "checklist-item-row" });

    const order = el("div", {
      className: "checklist-item-order",
      textContent: String(index + 1),
    });

    const input = el("input", {
      type: "text",
      className: "checklist-item-input",
      value: item.text,
      placeholder: "Enter checklist item...",
      oninput: (e) => {
        updateChecklistItem(checklist.id, item.id, e.target.value);
      },
    });

    const removeBtn = el("button", {
      className: "checklist-item-remove",
      type: "button",
      title: "Remove item",
      onclick: () => {
        if (confirm('Remove this item?')) {
          removeChecklistItem(checklist.id, item.id);
          renderChecklistItems();
        }
      },
    });
    // Use an elegant X for the remove button instead of text
    removeBtn.innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>`;

    // SVG icon for drag handle (lucide styling)
    const dragHandle = el("div", { className: "checklist-drag-handle", title: "Drag to reorder (Coming soon)" });
    dragHandle.innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="9" cy="12" r="1"></circle><circle cx="9" cy="5" r="1"></circle><circle cx="9" cy="19" r="1"></circle><circle cx="15" cy="12" r="1"></circle><circle cx="15" cy="5" r="1"></circle><circle cx="15" cy="19" r="1"></circle></svg>`;

    row.append(dragHandle, order, input, removeBtn);
    container.appendChild(row);
  });
}

// Event listeners for checklist tab
document.getElementById("checklistName")?.addEventListener("input", (e) => {
  const checklist = getChecklistById(activeChecklistTabId);
  if (checklist && !checklist.isLocked) {
    checklist.name = e.target.value;
    debouncedSaveChecklists();
    renderChecklistTabs();
  }
});

document.getElementById("resetChecklistBtn")?.addEventListener("click", () => {
  if (confirm("Reset this checklist to default items? This will remove any custom items.")) {
    resetChecklistToDefault(activeChecklistTabId);
    renderChecklistItems();
  }
});

document.getElementById("deleteChecklistBtn")?.addEventListener("click", () => {
  const checklist = getChecklistById(activeChecklistTabId);
  if (checklist && confirm(`Delete checklist "${checklist.name}"?`)) {
    deleteChecklist(checklist.id);
    activeChecklistTabId = checklistsDb.checklists[0]?.id || null;
    renderChecklistTabs();
  }
});

document.getElementById("addChecklistItemBtn")?.addEventListener("click", () => {
  const checklist = getChecklistById(activeChecklistTabId);
  if (checklist) {
    addChecklistItem(checklist.id, "New checklist item");
    renderChecklistItems();
    // Focus the new item
    setTimeout(() => {
      const inputs = document.querySelectorAll(".checklist-item-input");
      if (inputs.length > 0) {
        inputs[inputs.length - 1].focus();
        inputs[inputs.length - 1].select();
      }
    }, 50);
  }
});

// ===================== CHECKLIST MODAL FUNCTIONS =====================

let activeChecklistView = null;

// Checklist modal state
let checklistModalState = {
  appliedTabs: [], // { instanceId, checklistId, instanceIndex }
  activeInstanceId: null,
};

function getActiveWorkroomContext() {
  const project = db[activeChecklistProject];
  if (!project) return { project: null, deliverables: [], deliverable: null };
  const deliverables = getProjectDeliverables(project);
  const deliverableIndex = Number(activeChecklistDeliverable);
  const deliverable =
    Number.isInteger(deliverableIndex) && deliverableIndex >= 0
      ? deliverables[deliverableIndex] || null
      : null;
  return { project, deliverables, deliverable };
}

function getWorkroomDeliverableKeywordCorpus(deliverable) {
  if (!deliverable) return "";
  const taskText = Array.isArray(deliverable.tasks)
    ? deliverable.tasks
      .map((task) => {
        if (!task) return "";
        if (typeof task === "string") return task;
        return task.text || "";
      })
      .join(" ")
    : "";
  return [deliverable.name || "", deliverable.notes || "", taskText]
    .join(" ")
    .toLowerCase();
}

function inferWorkroomCadDiscipline(deliverable, candidates = []) {
  const options = candidates.filter((discipline) =>
    WORKROOM_CAD_DISCIPLINES.includes(discipline)
  );
  if (!options.length) return "Electrical";
  if (options.length === 1) return options[0];

  const corpus = getWorkroomDeliverableKeywordCorpus(deliverable);
  if (!corpus) return options[0];

  let bestDiscipline = options[0];
  let bestScore = -1;
  let tie = false;

  options.forEach((discipline) => {
    const patterns = WORKROOM_DISCIPLINE_KEYWORDS[discipline] || [];
    const score = patterns.reduce((total, pattern) => {
      const matches = corpus.match(pattern);
      return total + (matches ? matches.length : 0);
    }, 0);
    if (score > bestScore) {
      bestScore = score;
      bestDiscipline = discipline;
      tie = false;
    } else if (score === bestScore) {
      tie = true;
    }
  });

  if (bestScore <= 0 || tie) return options[0];
  return bestDiscipline;
}

function ensureWorkroomCadDiscipline(deliverable, disciplines = []) {
  const options = disciplines.length ? disciplines : getWorkroomAvailableDisciplines();
  if (!options.length) return "Electrical";
  if (!deliverable) return options[0];

  const existing = normalizeWorkroomCadDiscipline(deliverable.workroomCadDiscipline, "");
  if (existing && options.includes(existing)) return existing;

  const inferred = inferWorkroomCadDiscipline(deliverable, options);
  if (deliverable.workroomCadDiscipline !== inferred) {
    deliverable.workroomCadDiscipline = inferred;
    debouncedSave();
  }
  return inferred;
}

function renderWorkroomCadRoutingControl() {
  const container = document.getElementById("workroomCadRouting");
  if (!container) return;

  const { deliverable } = getActiveWorkroomContext();
  const disciplines = getWorkroomAvailableDisciplines();

  if (!deliverable || disciplines.length <= 1) {
    container.hidden = true;
    container.innerHTML = "";
    return;
  }

  const activeDiscipline = ensureWorkroomCadDiscipline(deliverable, disciplines);
  container.hidden = false;
  container.innerHTML = "";

  container.appendChild(
    el("p", {
      className: "workroom-cad-routing-label",
      textContent: "CAD Folder Discipline",
    })
  );

  const group = el("div", {
    className: "workroom-cad-routing-group",
    role: "group",
    "aria-label": "CAD folder discipline",
  });

  disciplines.forEach((discipline) => {
    group.appendChild(
      el("button", {
        type: "button",
        className: `workroom-cad-routing-btn ${discipline === activeDiscipline ? "active" : ""}`,
        textContent: discipline,
        title: `Route Workroom CAD tools to ${discipline} folder`,
        "aria-pressed": String(discipline === activeDiscipline),
        onclick: () => {
          if (deliverable.workroomCadDiscipline === discipline) return;
          deliverable.workroomCadDiscipline = discipline;
          save();
          renderWorkroomCadRoutingControl();
        },
      })
    );
  });

  container.appendChild(group);
}

function buildWorkroomCadLaunchContext() {
  const { project, deliverable } = getActiveWorkroomContext();
  const disciplines = getWorkroomAvailableDisciplines();
  const activeDiscipline = ensureWorkroomCadDiscipline(deliverable, disciplines);
  return {
    source: "workroom",
    projectPath: String(project?.path || "").trim(),
    discipline: activeDiscipline || disciplines[0] || "Electrical",
  };
}

function consumePendingCadLaunchContext() {
  const context = pendingCadLaunchContext;
  pendingCadLaunchContext = null;
  return context;
}

function initChecklistModalTabs(project, deliverableIndex) {
  const deliverables = getProjectDeliverables(project);
  const resolvedIndex = Number(deliverableIndex);
  const deliverable =
    Number.isInteger(resolvedIndex) && resolvedIndex >= 0
      ? deliverables[resolvedIndex] || null
      : null;

  checklistModalState.appliedTabs = [];
  if (!deliverable) {
    checklistModalState.activeInstanceId = null;
    return;
  }

  if (!Array.isArray(deliverable.appliedChecklists)) {
    deliverable.appliedChecklists = [];
  }

  if (deliverable.appliedChecklists.length === 0) {
    const defaultChecklist =
      checklistsDb.checklists.find((c) => c.isDefault) || checklistsDb.checklists[0];
    if (defaultChecklist) {
      deliverable.appliedChecklists.push({
        instanceId: generateChecklistInstanceId(),
        checklistId: defaultChecklist.id,
        completedItems: [],
        itemNotes: {},
      });
      save();
    }
  }

  checklistModalState.appliedTabs = deliverable.appliedChecklists.map(
    (instance, instanceIndex) => ({
      instanceId: instance.instanceId,
      checklistId: instance.checklistId,
      instanceIndex,
    })
  );

  const hasActive = checklistModalState.appliedTabs.some(
    (tab) => tab.instanceId === checklistModalState.activeInstanceId
  );
  if (!hasActive) {
    checklistModalState.activeInstanceId = checklistModalState.appliedTabs[0]?.instanceId || null;
  }
  activeChecklistView = checklistModalState.activeInstanceId;
}

function openChecklistModal(projectIndex) {
  const project = db[projectIndex];
  if (!project) return;

  activeChecklistProject = projectIndex;
  activeChecklistDeliverable = null;
  checklistModalState.activeInstanceId = null;
  activeChecklistView = null;
  activeWorkroomLeftTab = "tasks";

  populateChecklistDeliverableSelect(project);
  document.getElementById("checklistModal").showModal();
  renderWorkroomToolStatusFooter();
}

function populateChecklistDeliverableSelect(project) {
  const select = document.getElementById("checklistDeliverableSelect");
  if (!select) return;

  select.innerHTML = "";
  const deliverables = getProjectDeliverables(project);

  if (deliverables.length === 0) {
    select.innerHTML = '<option value="">No deliverables found</option>';
    select.disabled = true;
    activeChecklistDeliverable = -1;
    checklistModalState.appliedTabs = [];
    checklistModalState.activeInstanceId = null;
    renderProjectWorkroom();
    return;
  }

  select.disabled = false;
  deliverables.forEach((deliverable, index) => {
    select.appendChild(
      el("option", {
        value: String(index),
        textContent: `${deliverable.name || "Untitled"}${deliverable.due ? " (Due: " + humanDate(deliverable.due) + ")" : ""
          }`,
      })
    );
  });

  const primaryIndex = deliverables.findIndex((d) => d.id === project.overviewDeliverableId);
  activeChecklistDeliverable = primaryIndex >= 0 ? primaryIndex : 0;
  select.value = String(activeChecklistDeliverable);

  initChecklistModalTabs(project, activeChecklistDeliverable);
  renderProjectWorkroom();
}

function normalizeWorkroomLeftTab(tabName) {
  const normalized = String(tabName || "").trim().toLowerCase();
  if (normalized === "tasks" || normalized === "notes" || normalized === "checklist") {
    return normalized;
  }
  return "tasks";
}

function renderWorkroomLeftTabs() {
  const tabs = Array.from(document.querySelectorAll("[data-workroom-left-tab]"));
  const panes = Array.from(document.querySelectorAll("[data-workroom-left-pane]"));
  if (!tabs.length || !panes.length) return;

  activeWorkroomLeftTab = normalizeWorkroomLeftTab(activeWorkroomLeftTab);

  tabs.forEach((tabBtn) => {
    const tabName = normalizeWorkroomLeftTab(tabBtn.dataset.workroomLeftTab);
    const isActive = tabName === activeWorkroomLeftTab;
    tabBtn.classList.toggle("active", isActive);
    tabBtn.setAttribute("aria-selected", isActive ? "true" : "false");
    tabBtn.setAttribute("tabindex", isActive ? "0" : "-1");
    tabBtn.onclick = () => {
      const nextTab = normalizeWorkroomLeftTab(tabBtn.dataset.workroomLeftTab);
      if (nextTab === activeWorkroomLeftTab) return;
      activeWorkroomLeftTab = nextTab;
      renderWorkroomLeftTabs();
    };
  });

  panes.forEach((pane) => {
    const paneName = normalizeWorkroomLeftTab(pane.dataset.workroomLeftPane);
    pane.hidden = paneName !== activeWorkroomLeftTab;
    pane.setAttribute("aria-hidden", pane.hidden ? "true" : "false");
  });
}

function renderProjectWorkroom() {
  renderWorkroomLeftTabs();
  renderWorkroomChecklistPanel();
  renderWorkroomTasksPanel();
  renderWorkroomDeliverableNotesPanel();
  renderWorkroomGeneralNotesPanel();
  renderWorkroomToolsPanel();
}

function getToolTitle(toolId) {
  const card = document.getElementById(toolId);
  if (!card) return "Tool";
  return card.querySelector(".tool-card-header")?.textContent?.trim() || "Tool";
}

function renderWorkroomToolStatusFooter() {
  const statusEl = document.getElementById("workroomToolStatus");
  if (!statusEl) return;
  statusEl.classList.remove("is-running", "is-done", "is-error");
  const phase = workroomToolStatusState?.phase || "idle";
  if (phase === "running") statusEl.classList.add("is-running");
  if (phase === "done") statusEl.classList.add("is-done");
  if (phase === "error") statusEl.classList.add("is-error");
  statusEl.textContent = workroomToolStatusState?.message || "Ready.";
}

function setWorkroomToolStatus({ toolId = "", message = "", phase = "idle" } = {}) {
  const normalizedPhase =
    phase === "running" || phase === "done" || phase === "error"
      ? phase
      : "idle";
  const toolLabel = toolId ? getToolTitle(toolId) : "";
  let nextMessage = String(message || "").trim();
  if (!nextMessage) nextMessage = normalizedPhase === "done" ? "Done." : "Ready.";
  if (toolLabel && normalizedPhase !== "idle") {
    nextMessage = `${toolLabel}: ${nextMessage}`;
  }
  workroomToolStatusState = {
    toolId,
    label: toolLabel,
    message: nextMessage,
    phase: normalizedPhase,
  };
  renderWorkroomToolStatusFooter();
}

function resetWorkroomToolStatus() {
  workroomToolStatusState = {
    toolId: "",
    label: "",
    message: "Ready.",
    phase: "idle",
  };
  renderWorkroomToolStatusFooter();
}

function getWorkroomVisibleTools() {
  return Array.from(
    document.querySelectorAll("#tools-panel .tool-card:not([hidden])")
  )
    .map((card) => {
      const id = card.id;
      if (!id || WORKROOM_HIDDEN_TOOL_IDS.has(id)) return null;
      const iconHost = card.querySelector(".tool-icon");
      const iconMarkup = iconHost?.innerHTML?.trim() || "";
      const hasGraphicIcon = !!iconHost?.querySelector("svg, img");
      return {
        id,
        iconMarkup,
        hasGraphicIcon,
        iconTextFallback: iconHost?.textContent?.trim() || "TOOL",
        title:
          card.querySelector(".tool-card-header")?.textContent?.trim() ||
          "Unnamed Tool",
      };
    })
    .filter(Boolean);
}

function triggerWorkroomTool(toolId) {
  if (!toolId) return;
  const sourceCard = document.getElementById(toolId);
  if (!sourceCard || sourceCard.hidden) {
    toast("Selected tool is unavailable.");
    return;
  }
  pendingCadLaunchContext = null;
  if (WORKROOM_LAUNCH_CONTEXT_TOOL_IDS.has(toolId)) {
    pendingCadLaunchContext = buildWorkroomCadLaunchContext();
  }
  setWorkroomToolStatus({ toolId, message: "Starting...", phase: "running" });
  sourceCard.click();
}

function renderWorkroomToolsPanel() {
  const toolsList = document.getElementById("workroomToolsList");
  if (!toolsList) return;
  renderWorkroomCadRoutingControl();

  const tools = getWorkroomVisibleTools();
  toolsList.innerHTML = "";

  if (!tools.length) {
    toolsList.appendChild(
      el("div", {
        className: "workroom-tool-empty",
        textContent: "No tools available.",
      })
    );
    renderWorkroomToolStatusFooter();
    return;
  }

  tools.forEach((tool) => {
    const item = el("div", {
      className: "workroom-tool-item",
      "data-tool-id": tool.id,
    });
    const iconEl = el("span", {
      className: "workroom-tool-icon",
      "aria-hidden": "true",
    });
    if (tool.hasGraphicIcon && tool.iconMarkup) {
      iconEl.innerHTML = tool.iconMarkup;
    } else {
      iconEl.textContent = tool.iconTextFallback;
    }
    item.appendChild(iconEl);
    item.appendChild(
      el("span", {
        className: "workroom-tool-text",
        textContent: tool.title,
        title: tool.title,
      })
    );
    item.appendChild(
      el("button", {
        className: "btn tiny workroom-tool-run-btn",
        type: "button",
        textContent: "Run",
        title: `Run ${tool.title}`,
        "aria-label": `Run ${tool.title}`,
        onclick: (e) => {
          e.stopPropagation();
          triggerWorkroomTool(tool.id);
        },
      })
    );
    toolsList.appendChild(item);
  });
  renderWorkroomToolStatusFooter();
}

function renderWorkroomChecklistPanel() {
  const itemsContainer = document.getElementById("workroomChecklistItems");
  if (!itemsContainer) return;

  renderWorkroomChecklistTabs();

  const { deliverable } = getActiveWorkroomContext();
  if (!deliverable) {
    itemsContainer.innerHTML = "<div class='chk-empty-state'>Select a deliverable to view checklists.</div>";
    setWorkroomChecklistProgress(0, 0);
    return;
  }

  if (!checklistModalState.appliedTabs.length) {
    itemsContainer.innerHTML =
      "<div class='chk-empty-state'>No checklists applied. Use the add checklist dropdown.</div>";
    setWorkroomChecklistProgress(0, 0);
    return;
  }

  const activeTab =
    checklistModalState.appliedTabs.find(
      (tab) => tab.instanceId === checklistModalState.activeInstanceId
    ) || checklistModalState.appliedTabs[0];

  checklistModalState.activeInstanceId = activeTab.instanceId;
  activeChecklistView = activeTab.instanceId;

  const instance = deliverable.appliedChecklists?.[activeTab.instanceIndex];
  const checklist = getChecklistById(activeTab.checklistId);
  if (!instance || !checklist) {
    itemsContainer.innerHTML = "<div class='chk-empty-state'>Checklist instance not found.</div>";
    setWorkroomChecklistProgress(0, 0);
    return;
  }

  renderWorkroomChecklistItems(itemsContainer, instance, checklist, activeTab.instanceIndex);
}

function renderWorkroomChecklistTabs() {
  const tabsContainer = document.getElementById("workroomChecklistTabs");
  const addSelect = document.getElementById("workroomChecklistAddSelect");
  if (!tabsContainer || !addSelect) return;

  tabsContainer.innerHTML = "";

  const { deliverable } = getActiveWorkroomContext();
  if (!deliverable) {
    addSelect.disabled = true;
    addSelect.innerHTML = '<option value="">+ Add Checklist...</option>';
    return;
  }

  checklistModalState.appliedTabs.forEach((tab, tabIndex) => {
    const checklist = getChecklistById(tab.checklistId);
    const instance = deliverable.appliedChecklists?.[tab.instanceIndex];
    const completedCount = instance?.completedItems?.length || 0;
    const totalCount = checklist?.items?.length || 0;
    const isActive = tab.instanceId === checklistModalState.activeInstanceId;

    const tabBtn = el("button", {
      className: `workroom-checklist-tab ${isActive ? "active" : ""}`,
      type: "button",
      onclick: () => {
        checklistModalState.activeInstanceId = tab.instanceId;
        activeChecklistView = tab.instanceId;
        renderWorkroomChecklistPanel();
      },
    });
    tabBtn.appendChild(
      el("span", {
        className: "workroom-checklist-tab-label",
        textContent: checklist?.name || "Checklist",
      })
    );
    tabBtn.appendChild(
      el("span", {
        className: "workroom-checklist-tab-badge",
        textContent: `${completedCount}/${totalCount}`,
      })
    );

    const removeBtn = el("button", {
      className: "workroom-checklist-remove icon-btn mini",
      type: "button",
      title: "Remove checklist",
      "aria-label": "Remove checklist",
      onclick: (e) => {
        e.stopPropagation();
        if (confirm(`Remove "${checklist?.name || "checklist"}" from this deliverable?`)) {
          removeChecklistInstance(tabIndex);
        }
      },
    });
    removeBtn.textContent = "x";
    tabBtn.appendChild(removeBtn);
    tabsContainer.appendChild(tabBtn);
  });

  populateWorkroomChecklistAddSelect();
  addSelect.disabled = false;
  addSelect.onchange = (e) => {
    const checklistId = e.target.value;
    if (!checklistId) return;
    addWorkroomChecklistInstance(checklistId);
    e.target.value = "";
  };
}

function populateWorkroomChecklistAddSelect() {
  const select = document.getElementById("workroomChecklistAddSelect");
  if (!select) return;

  const currentIds = checklistModalState.appliedTabs.map((tab) => tab.checklistId);
  select.innerHTML = '<option value="">+ Add Checklist...</option>';

  checklistsDb.checklists.forEach((checklist) => {
    if (currentIds.includes(checklist.id)) return;
    select.appendChild(
      el("option", {
        value: checklist.id,
        textContent: checklist.name,
      })
    );
  });
}

function addWorkroomChecklistInstance(checklistId) {
  const { project, deliverable } = getActiveWorkroomContext();
  if (!project || !deliverable) return;

  if (!Array.isArray(deliverable.appliedChecklists)) {
    deliverable.appliedChecklists = [];
  }
  if (deliverable.appliedChecklists.some((instance) => instance.checklistId === checklistId)) {
    toast("Checklist already added to this deliverable.");
    return;
  }

  const instance = {
    instanceId: generateChecklistInstanceId(),
    checklistId,
    completedItems: [],
    itemNotes: {},
  };
  deliverable.appliedChecklists.push(instance);
  save();

  initChecklistModalTabs(project, activeChecklistDeliverable);
  checklistModalState.activeInstanceId = instance.instanceId;
  activeChecklistView = instance.instanceId;
  renderWorkroomChecklistPanel();
}

function removeChecklistInstance(tabIndex) {
  const tab = checklistModalState.appliedTabs[tabIndex];
  if (!tab) return;

  const { project, deliverable } = getActiveWorkroomContext();
  if (!project || !deliverable || !Array.isArray(deliverable.appliedChecklists)) return;

  deliverable.appliedChecklists.splice(tab.instanceIndex, 1);
  save();

  initChecklistModalTabs(project, activeChecklistDeliverable);
  renderWorkroomChecklistPanel();
}

function renderWorkroomChecklistItems(container, instance, checklist, instanceIndex) {
  container.innerHTML = "";

  const completedItems = Array.isArray(instance.completedItems) ? instance.completedItems : [];
  const checklistItems = Array.isArray(checklist.items) ? checklist.items : [];

  if (!checklistItems.length) {
    container.innerHTML = "<div class='chk-empty-state'>This checklist has no items.</div>";
    setWorkroomChecklistProgress(0, 0);
    return;
  }

  checklistItems.forEach((item) => {
    const isCompleted = completedItems.includes(item.id);
    const row = el("div", {
      className: `checklist-modal-item ${isCompleted ? "checked" : ""}`,
      onclick: () => toggleWorkroomChecklistItem(instanceIndex, item.id),
    });

    const checkLabel = el("label", { className: "custom-check" });
    checkLabel.addEventListener("click", (e) => e.stopPropagation());
    const checkbox = el("input", {
      type: "checkbox",
      checked: isCompleted,
      onchange: () => toggleWorkroomChecklistItem(instanceIndex, item.id),
    });
    const checkmarkSpan = el("span", { className: "checkmark" });
    checkLabel.append(checkbox, checkmarkSpan);

    row.append(
      checkLabel,
      el("span", {
        className: "checklist-modal-item-text",
        textContent: item.text,
      })
    );
    container.appendChild(row);
  });

  setWorkroomChecklistProgress(completedItems.length, checklistItems.length);
}

function toggleWorkroomChecklistItem(instanceIndex, itemId) {
  const { deliverable } = getActiveWorkroomContext();
  const instance = deliverable?.appliedChecklists?.[instanceIndex];
  if (!instance) return;

  if (!Array.isArray(instance.completedItems)) {
    instance.completedItems = [];
  }

  const idx = instance.completedItems.indexOf(itemId);
  if (idx >= 0) {
    instance.completedItems.splice(idx, 1);
  } else {
    instance.completedItems.push(itemId);
  }

  save();
  renderWorkroomChecklistPanel();
}

function setWorkroomChecklistProgress(done, total) {
  const countEl = document.getElementById("workroomChecklistProgressCount");
  const fillEl = document.getElementById("workroomChecklistProgressFill");
  if (countEl) countEl.textContent = `${done}/${total}`;
  if (fillEl) fillEl.style.width = total > 0 ? `${(done / total) * 100}%` : "0%";
}

function renderWorkroomTasksPanel() {
  const addTaskBtn = document.getElementById("workroomAddTaskBtn");
  const container = document.getElementById("workroomTasksContainer");
  if (!addTaskBtn || !container) return;

  const { deliverable } = getActiveWorkroomContext();
  if (!deliverable) {
    addTaskBtn.disabled = true;
    addTaskBtn.onclick = null;
    container.innerHTML = "<div class='chk-empty-state'>Select a deliverable to manage tasks.</div>";
    setWorkroomTaskProgress(0, 0);
    return;
  }

  addTaskBtn.disabled = false;
  addTaskBtn.onclick = () => {
    if (!Array.isArray(deliverable.tasks)) {
      deliverable.tasks = [];
    }
    deliverable.tasks.push({ text: "New task", done: false, link: "", link2: "" });
    save();
    renderWorkroomTasksPanel();
    setTimeout(() => {
      const inputs = container.querySelectorAll(".checklist-task-text-input");
      const last = inputs[inputs.length - 1];
      if (last) {
        last.focus();
        last.select();
      }
    }, 20);
  };

  if (!Array.isArray(deliverable.tasks) || deliverable.tasks.length === 0) {
    container.innerHTML = "<div class='chk-empty-state'>No tasks yet.</div>";
    setWorkroomTaskProgress(0, 0);
    return;
  }

  container.innerHTML = "";
  deliverable.tasks.forEach((task, taskIndex) => {
    const row = el("div", {
      className: `checklist-task-row ${task.done ? "checklist-task-done" : ""}`,
    });

    const checkLabel = el("label", { className: "custom-check" });
    const checkbox = el("input", {
      type: "checkbox",
      checked: !!task.done,
      onchange: () => {
        task.done = checkbox.checked;
        save();
        renderWorkroomTasksPanel();
      },
    });
    checkLabel.append(checkbox, el("span", { className: "checkmark" }));

    const textInput = el("input", {
      type: "text",
      className: "checklist-task-text-input",
      value: task.text || "",
      placeholder: "Task description...",
      oninput: (e) => {
        task.text = e.target.value;
        debouncedSave();
      },
      onkeydown: (e) => {
        if (e.key === "Enter") {
          deliverable.tasks.splice(taskIndex + 1, 0, { text: "", done: false, link: "", link2: "" });
          save();
          renderWorkroomTasksPanel();
          setTimeout(() => {
            const inputs = container.querySelectorAll(".checklist-task-text-input");
            if (inputs[taskIndex + 1]) inputs[taskIndex + 1].focus();
          }, 20);
        } else if (e.key === "Backspace" && !e.target.value) {
          deliverable.tasks.splice(taskIndex, 1);
          save();
          renderWorkroomTasksPanel();
        }
      },
    });

    const removeBtn = el("button", {
      className: "checklist-task-remove",
      type: "button",
      textContent: "x",
      onclick: () => {
        deliverable.tasks.splice(taskIndex, 1);
        save();
        renderWorkroomTasksPanel();
      },
    });

    row.append(checkLabel, textInput, removeBtn);
    container.appendChild(row);
  });

  const completed = deliverable.tasks.filter((task) => task.done).length;
  setWorkroomTaskProgress(completed, deliverable.tasks.length);
}

function setWorkroomTaskProgress(done, total) {
  const countEl = document.getElementById("workroomTaskProgressCount");
  const fillEl = document.getElementById("workroomTaskProgressFill");
  if (countEl) countEl.textContent = `${done}/${total}`;
  if (fillEl) fillEl.style.width = total > 0 ? `${(done / total) * 100}%` : "0%";
}

function renderWorkroomDeliverableNotesPanel() {
  const notesTextarea = document.getElementById("workroomDeliverableNotesTextarea");
  if (!notesTextarea) return;

  const { deliverable } = getActiveWorkroomContext();
  if (!deliverable) {
    notesTextarea.value = "";
    notesTextarea.disabled = true;
    notesTextarea.placeholder = "Select a deliverable to view notes.";
    notesTextarea.oninput = null;
    notesTextarea.onblur = null;
    return;
  }

  notesTextarea.disabled = false;
  notesTextarea.placeholder = "Add notes for this deliverable...";
  notesTextarea.value = deliverable.notes || "";
  notesTextarea.oninput = () => {
    deliverable.notes = notesTextarea.value;
    debouncedSave();
    autoResizeTextarea(notesTextarea);
  };
  notesTextarea.onblur = () => {
    deliverable.notes = notesTextarea.value;
    save();
  };
  requestAnimationFrame(() => autoResizeTextarea(notesTextarea));
}

function ensureActiveNoteTabSelection() {
  if (!Array.isArray(noteTabs) || noteTabs.length === 0) {
    noteTabs = ["General"];
  }
  if (!activeNoteTab || !noteTabs.includes(activeNoteTab)) {
    activeNoteTab = noteTabs[0];
  }
  if (!notesDb || typeof notesDb !== "object") {
    notesDb = {};
  }
  noteTabs.forEach((tab) => {
    if (typeof notesDb[tab] !== "string") {
      notesDb[tab] = notesDb[tab] || "";
    }
  });
}

function syncMainNotesFromWorkroom() {
  updateActiveNoteTextarea();
  const mainNotesTextarea = document.getElementById("notesTextarea");
  if (mainNotesTextarea && activeNoteTab) {
    mainNotesTextarea.value = notesDb[activeNoteTab] || "";
  }
}

function renderWorkroomGeneralNotesPanel() {
  const tabsContainer = document.getElementById("workroomNotesTabs");
  const addBtn = document.getElementById("workroomAddNotePageBtn");
  const notesTextarea = document.getElementById("workroomGeneralNotesTextarea");
  const searchInput = document.getElementById("workroomNotesSearchInput");
  const searchBtn = document.getElementById("workroomNotesSearchBtn");
  const resultsContainer = document.getElementById("workroomNotesSearchResults");
  if (!tabsContainer || !addBtn || !notesTextarea || !searchInput || !searchBtn || !resultsContainer)
    return;

  ensureActiveNoteTabSelection();
  tabsContainer.innerHTML = "";

  noteTabs.forEach((tabName) => {
    const tabBtn = el("button", {
      className: `workroom-note-tab ${tabName === activeNoteTab ? "active" : ""}`,
      type: "button",
      onclick: () => {
        activeNoteTab = tabName;
        renderWorkroomGeneralNotesPanel();
        renderNoteTabs();
        renderNoteSearchResults();
      },
    });
    tabBtn.appendChild(
      el("span", {
        className: "workroom-note-tab-label",
        textContent: tabName,
      })
    );

    const removeBtn = el("button", {
      className: "workroom-note-tab-remove icon-btn mini",
      type: "button",
      title: "Delete page",
      "aria-label": "Delete page",
      onclick: (e) => {
        e.stopPropagation();
        if (confirm(`Permanently delete page "${tabName}"?`)) {
          const idx = noteTabs.indexOf(tabName);
          if (idx > -1) {
            noteTabs.splice(idx, 1);
            delete notesDb[tabName];
            activeNoteTab = noteTabs[Math.max(0, idx - 1)] || null;
            ensureActiveNoteTabSelection();
            saveNotes();
            renderWorkroomGeneralNotesPanel();
            renderNoteTabs();
            renderNoteSearchResults();
          }
        }
      },
    });
    removeBtn.textContent = "x";
    tabBtn.appendChild(removeBtn);
    tabsContainer.appendChild(tabBtn);
  });

  addBtn.onclick = () => {
    const name = prompt("Enter name for new page:");
    if (!name || !name.trim()) return;
    const trimmed = name.trim();
    if (noteTabs.includes(trimmed)) {
      toast("Page name already exists.");
      return;
    }
    noteTabs.push(trimmed);
    notesDb[trimmed] = "";
    activeNoteTab = trimmed;
    saveNotes();
    renderWorkroomGeneralNotesPanel();
    renderNoteTabs();
    renderNoteSearchResults();
  };

  searchBtn.onclick = () => {
    renderWorkroomNoteSearchResults();
  };
  searchInput.onkeydown = (e) => {
    if (e.key !== "Enter") return;
    e.preventDefault();
    renderWorkroomNoteSearchResults();
  };

  notesTextarea.disabled = !activeNoteTab;
  notesTextarea.placeholder = activeNoteTab
    ? `Enter notes for ${activeNoteTab}...`
    : "Create a page to begin.";
  notesTextarea.value = activeNoteTab ? notesDb[activeNoteTab] || "" : "";

  notesTextarea.oninput = () => {
    if (!activeNoteTab) return;
    notesDb[activeNoteTab] = notesTextarea.value;
    debouncedSaveNotes();
    syncMainNotesFromWorkroom();
  };
  notesTextarea.onblur = () => {
    if (!activeNoteTab) return;
    notesDb[activeNoteTab] = notesTextarea.value;
    saveNotes();
    syncMainNotesFromWorkroom();
  };

  if ((searchInput.value || "").trim()) {
    renderWorkroomNoteSearchResults();
  } else {
    resultsContainer.innerHTML = "";
  }
}

document.getElementById("checklistDeliverableSelect")?.addEventListener("change", (e) => {
  const project = db[activeChecklistProject];
  if (!project) return;

  activeChecklistDeliverable = Number(e.target.value);
  initChecklistModalTabs(project, activeChecklistDeliverable);
  renderProjectWorkroom();
});

// Close modal handler
document.getElementById("checklistModal")?.addEventListener("close", () => {
  activeChecklistProject = null;
  activeChecklistDeliverable = null;
  activeChecklistView = null;
  activeWorkroomLeftTab = "tasks";
  checklistModalState.appliedTabs = [];
  checklistModalState.activeInstanceId = null;
  pendingCadLaunchContext = null;
  const workroomToolsSettingsDlg = document.getElementById(
    "workroomToolsSettingsDlg"
  );
  if (workroomToolsSettingsDlg?.open) {
    workroomToolsSettingsDlg.close();
  }
  resetWorkroomToolStatus();
  const workroomSearchInput = document.getElementById("workroomNotesSearchInput");
  const workroomResults = document.getElementById("workroomNotesSearchResults");
  if (workroomSearchInput) workroomSearchInput.value = "";
  if (workroomResults) workroomResults.innerHTML = "";
});

// Close button handler
document.getElementById("checklistModalCloseBtn")?.addEventListener("click", (e) => {
  e.stopPropagation();
  document.getElementById("checklistModal").close();
});

// Make functions globally available
window.openChecklistModal = openChecklistModal;
window.switchChecklistView = () => { };
window.__aciesAutomation = {
  openWorkroom(projectIndex = 0) {
    const idx = Number.isInteger(Number(projectIndex)) ? Number(projectIndex) : 0;
    openChecklistModal(idx);
    return this.getWorkroomState();
  },
  closeWorkroom() {
    const modal = document.getElementById("checklistModal");
    if (modal?.open) modal.close();
    return this.getWorkroomState();
  },
  runWorkroomTool(toolId) {
    triggerWorkroomTool(toolId);
    return this.getToolState(toolId);
  },
  getWorkroomState() {
    const modal = document.getElementById("checklistModal");
    const deliverableSelect = document.getElementById("checklistDeliverableSelect");
    return {
      modalOpen: !!modal?.open,
      activeProjectIndex: activeChecklistProject,
      activeDeliverableIndex: activeChecklistDeliverable,
      deliverableOptions: deliverableSelect?.options?.length || 0,
    };
  },
  getToolState(toolId) {
    const card = document.getElementById(toolId);
    const statusEl = card?.querySelector(".tool-card-status");
    const statusText = (statusEl?.textContent || "").trim();
    const running = !!card?.classList.contains("running");
    let phase = "idle";
    if (statusText.toLowerCase().startsWith("error")) {
      phase = "error";
    } else if (statusText.toLowerCase().includes("done")) {
      phase = "done";
    } else if (running) {
      phase = "running";
    }
    return {
      toolId,
      exists: !!card,
      running,
      phase,
      statusText,
      workroom: this.getWorkroomState(),
    };
  },
  simulateAiResult(aiData = {}) {
    return handleAiProjectResult(aiData || {});
  },
  getAiNoMatchDialogState() {
    return getAiNoMatchDialogState();
  },
  openAiNoMatchSearch() {
    return openAiNoMatchSearch();
  },
  setAiNoMatchSearch(query = "") {
    return setAiNoMatchSearchQuery(query);
  },
  selectAiNoMatchProject(indexOrPosition) {
    const selected = selectAiNoMatchProject(indexOrPosition);
    return { selected, state: getAiNoMatchDialogState() };
  },
  confirmAiNoMatchAddToProject() {
    const added = confirmAiNoMatchAddToProject();
    return { added, state: getAiNoMatchDialogState(), edit: this.getEditDialogState() };
  },
  chooseAiNoMatchCreateNew() {
    const rawAiData = aiNoMatchState.rawAiData || {};
    closeAiNoMatchDialog();
    openAiCreateNewProject(rawAiData);
    return this.getEditDialogState();
  },
  closeAiNoMatchDialog() {
    closeAiNoMatchDialog();
    return getAiNoMatchDialogState();
  },
  getProjectSummary(projectIndex) {
    const idx = Number(projectIndex);
    if (!Number.isInteger(idx) || idx < 0 || !db[idx]) {
      return { exists: false, index: idx };
    }
    const project = normalizeProject(db[idx]);
    return {
      exists: true,
      index: idx,
      id: project.id || "",
      name: project.name || "",
      nick: project.nick || "",
      deliverableCount: getProjectDeliverables(project).length,
      deliverables: getProjectDeliverables(project).map((deliverable) => ({
        id: deliverable.id || "",
        name: deliverable.name || "",
        due: deliverable.due || "",
        taskCount: Array.isArray(deliverable.tasks) ? deliverable.tasks.length : 0,
      })),
    };
  },
  getEditDialogState() {
    const dlg = document.getElementById("editDlg");
    return {
      open: !!dlg?.open,
      editIndex,
      title: document.getElementById("dlgTitle")?.textContent || "",
      saveButtonText: document.getElementById("btnSaveProject")?.textContent || "",
      projectId: document.getElementById("f_id")?.value || "",
      projectName: document.getElementById("f_name")?.value || "",
      projectNick: document.getElementById("f_nick")?.value || "",
      projectPath: document.getElementById("f_path")?.value || "",
      deliverableCards:
        document.querySelectorAll("#deliverableList .deliverable-card-new").length ||
        0,
    };
  },
  commitEditDialog() {
    const button = document.getElementById("btnSaveProject");
    if (button) button.click();
    return this.getEditDialogState();
  },
  cancelEditDialog() {
    const dlg = document.getElementById("editDlg");
    if (dlg?.open) closeDlg("editDlg");
    return this.getEditDialogState();
  },
};
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
  const nextMessage = String(message || "").trim();
  let footerPhase = "running";
  let footerMessage = nextMessage || "Running...";

  if (nextMessage.startsWith("ERROR:")) {
    footerPhase = "error";
    footerMessage = nextMessage.substring(6).trim() || "Error.";
  } else if (nextMessage === "DONE") {
    footerPhase = "done";
    footerMessage = "Done.";
  }

  const checklistModal = document.getElementById("checklistModal");
  if (checklistModal?.open) {
    setWorkroomToolStatus({
      toolId,
      message: footerMessage,
      phase: footerPhase,
    });
  }

  if (!statusEl) return;
  statusEl.classList.remove("error");
  if (toolId === "toolCleanDwgs" && abortBtn) {
    if (nextMessage && nextMessage !== "DONE" && !nextMessage.startsWith("ERROR:")) {
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
  if (nextMessage.startsWith("ERROR:")) {
    statusEl.textContent = nextMessage.substring(6).trim();
    statusEl.classList.add("error");
    card.classList.add("running");
    setTimeout(() => {
      card.classList.remove("running");
      if (abortBtn) abortBtn.style.display = "none";
    }, 5000);
  } else if (nextMessage === "DONE") {
    statusEl.textContent = "Done!";
    setTimeout(() => {
      card.classList.remove("running");
      if (abortBtn) abortBtn.style.display = "none";
    }, 2000);
  } else {
    statusEl.textContent = nextMessage;
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
  panels: [],
  activePanelId: "",
  nextPanelNumber: 1,
  outputMode: "new",
  newOutputExtension: "xlsx",
  newOutputPath: "",
  existingOutputPath: "",
  running: false,
};

function generateCircuitBreakerPanelId() {
  return `panel_${Math.random().toString(36).slice(2, 10)}`;
}

function createCircuitBreakerPanel() {
  const number = circuitBreakerState.nextPanelNumber++;
  return {
    id: generateCircuitBreakerPanelId(),
    label: `Panel ${number}`,
    panelName: "",
    breakerPath: "",
    directoryPath: "",
    breakerFile: null,
    directoryFile: null,
  };
}

function ensureCircuitBreakerPanels() {
  if (Array.isArray(circuitBreakerState.panels) && circuitBreakerState.panels.length) {
    return;
  }
  const initialPanel = createCircuitBreakerPanel();
  circuitBreakerState.panels = [initialPanel];
  circuitBreakerState.activePanelId = initialPanel.id;
}

function getActiveCircuitBreakerPanel() {
  ensureCircuitBreakerPanels();
  const active =
    circuitBreakerState.panels.find(
      (panel) => panel.id === circuitBreakerState.activePanelId
    ) || circuitBreakerState.panels[0];
  if (active && active.id !== circuitBreakerState.activePanelId) {
    circuitBreakerState.activePanelId = active.id;
  }
  return active || null;
}

function getCircuitBreakerPanelDisplayName(panel) {
  if (!panel) return "Panel";
  const customName = String(panel.panelName || "").trim();
  return customName || panel.label || "Panel";
}

function isCircuitBreakerPanelReady(panel) {
  if (!panel) return false;
  return Boolean(
    (panel.breakerPath || panel.breakerFile) &&
    (panel.directoryPath || panel.directoryFile)
  );
}

function getCircuitBreakerReadyCounts() {
  ensureCircuitBreakerPanels();
  const total = circuitBreakerState.panels.length;
  const ready = circuitBreakerState.panels.filter((panel) =>
    isCircuitBreakerPanelReady(panel)
  ).length;
  return { ready, total };
}

function setActiveCircuitBreakerPanel(panelId) {
  if (circuitBreakerState.running) return;
  const panel = circuitBreakerState.panels.find((item) => item.id === panelId);
  if (!panel) return;
  circuitBreakerState.activePanelId = panel.id;
  updateCircuitBreakerUi();
}

function addCircuitBreakerPanel() {
  if (circuitBreakerState.running) return;
  ensureCircuitBreakerPanels();
  const panel = createCircuitBreakerPanel();
  circuitBreakerState.panels.push(panel);
  circuitBreakerState.activePanelId = panel.id;
  updateCircuitBreakerUi();
}

function removeCircuitBreakerPanel(panelId) {
  if (circuitBreakerState.running) return;
  ensureCircuitBreakerPanels();
  if (circuitBreakerState.panels.length <= 1) return;
  const index = circuitBreakerState.panels.findIndex(
    (panel) => panel.id === panelId
  );
  if (index < 0) return;
  circuitBreakerState.panels.splice(index, 1);
  if (circuitBreakerState.activePanelId === panelId) {
    const nextPanel =
      circuitBreakerState.panels[Math.max(0, index - 1)] ||
      circuitBreakerState.panels[0] ||
      null;
    circuitBreakerState.activePanelId = nextPanel ? nextPanel.id : "";
  }
  updateCircuitBreakerUi();
}

function getCircuitBreakerOutputPath() {
  return circuitBreakerState.outputMode === "existing"
    ? circuitBreakerState.existingOutputPath
    : circuitBreakerState.newOutputPath;
}

function normalizeCircuitBreakerOutputExtension(value) {
  const normalized = String(value || "")
    .trim()
    .toLowerCase()
    .replace(/^\./, "");
  if (normalized === "xls" || normalized === "xlsx") {
    return normalized;
  }
  return "xlsx";
}

function getCircuitBreakerOutputExtensionFromPath(path) {
  const ext = String(path || "")
    .trim()
    .split(".")
    .pop()
    ?.toLowerCase();
  if (ext === "xls" || ext === "xlsx") {
    return ext;
  }
  return "";
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

function renderCircuitBreakerPanelTabs() {
  const tabsContainer = document.getElementById("cbPanelTabs");
  if (!tabsContainer) return;
  ensureCircuitBreakerPanels();
  tabsContainer.innerHTML = "";
  const canRemove = circuitBreakerState.panels.length > 1 && !circuitBreakerState.running;
  circuitBreakerState.panels.forEach((panel) => {
    const isActive = panel.id === circuitBreakerState.activePanelId;
    const tab = el("div", {
      className: `cb-panel-tab ${isActive ? "is-active" : ""}`,
      role: "presentation",
    });

    const tabBtn = el("button", {
      className: "cb-panel-tab-btn",
      type: "button",
      title: getCircuitBreakerPanelDisplayName(panel),
      onclick: () => setActiveCircuitBreakerPanel(panel.id),
    });
    tabBtn.setAttribute("role", "tab");
    tabBtn.setAttribute("aria-selected", isActive ? "true" : "false");
    tabBtn.setAttribute("tabindex", isActive ? "0" : "-1");
    tabBtn.appendChild(
      el("span", {
        className: "cb-panel-tab-label",
        textContent: getCircuitBreakerPanelDisplayName(panel),
      })
    );
    tabBtn.appendChild(
      el("span", {
        className: "cb-panel-tab-state",
        textContent: isCircuitBreakerPanelReady(panel) ? "Ready" : "Pending",
      })
    );

    const removeBtn = el("button", {
      className: "cb-panel-tab-remove",
      type: "button",
      textContent: "x",
      title: `Remove ${panel.label || "panel"}`,
      disabled: !canRemove,
      onclick: (event) => {
        event.stopPropagation();
        removeCircuitBreakerPanel(panel.id);
      },
    });

    tab.appendChild(tabBtn);
    tab.appendChild(removeBtn);
    tabsContainer.appendChild(tab);
  });
}

function updateCircuitBreakerUi() {
  ensureCircuitBreakerPanels();
  const activePanel = getActiveCircuitBreakerPanel();
  const breakerFile = document.getElementById("cbBreakerFile");
  const directoryFile = document.getElementById("cbDirectoryFile");
  const newRow = document.getElementById("cbNewScheduleRow");
  const newFormatRow = document.getElementById("cbNewScheduleFormatRow");
  const newFormatSelect = document.getElementById("cbNewScheduleFormat");
  const existingRow = document.getElementById("cbExistingScheduleRow");
  const newPath = document.getElementById("cbNewSchedulePath");
  const existingPath = document.getElementById("cbExistingSchedulePath");
  const browseNewBtn = document.getElementById("cbBrowseNewScheduleBtn");
  const clearNewBtn = document.getElementById("cbClearNewScheduleBtn");
  const browseExistingBtn = document.getElementById("cbBrowseExistingScheduleBtn");
  const clearExistingBtn = document.getElementById("cbClearExistingScheduleBtn");
  const addPanelBtn = document.getElementById("cbAddPanelTabBtn");
  const panelNameInput = document.getElementById("cbPanelName");
  const runBtn = document.getElementById("cbRunPanelScheduleBtn");
  const modeNew = document.getElementById("cbOutputModeNew");
  const modeExisting = document.getElementById("cbOutputModeExisting");
  const breakerDrop = document.getElementById("cbBreakerDrop");
  const directoryDrop = document.getElementById("cbDirectoryDrop");
  const runningOverlay = document.getElementById("cbRunningOverlay");

  renderCircuitBreakerPanelTabs();

  if (breakerFile) {
    const label = getCircuitBreakerFileLabel(
      activePanel?.breakerPath,
      activePanel?.breakerFile
    );
    breakerFile.textContent = label;
    breakerFile.dataset.empty =
      activePanel && (activePanel.breakerPath || activePanel.breakerFile)
        ? "false"
        : "true";
  }
  if (directoryFile) {
    const label = getCircuitBreakerFileLabel(
      activePanel?.directoryPath,
      activePanel?.directoryFile
    );
    directoryFile.textContent = label;
    directoryFile.dataset.empty =
      activePanel && (activePanel.directoryPath || activePanel.directoryFile)
        ? "false"
        : "true";
  }
  if (newPath) {
    const label = circuitBreakerState.newOutputPath || "No file selected";
    newPath.textContent = label;
    newPath.title = label;
    newPath.dataset.empty = circuitBreakerState.newOutputPath ? "false" : "true";
  }
  if (existingPath) {
    const label = circuitBreakerState.existingOutputPath || "No file selected";
    existingPath.textContent = label;
    existingPath.title = label;
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
  if (newFormatRow) newFormatRow.hidden = circuitBreakerState.outputMode !== "new";
  if (existingRow) existingRow.hidden = circuitBreakerState.outputMode !== "existing";

  if (newFormatSelect) {
    const normalizedExtension = normalizeCircuitBreakerOutputExtension(
      circuitBreakerState.newOutputExtension
    );
    if (newFormatSelect.value !== normalizedExtension) {
      newFormatSelect.value = normalizedExtension;
    }
    newFormatSelect.disabled = circuitBreakerState.running;
  }

  if (panelNameInput && panelNameInput.value !== (activePanel?.panelName || "")) {
    panelNameInput.value = activePanel?.panelName || "";
  }

  const outputReady = Boolean(getCircuitBreakerOutputPath());
  const counts = getCircuitBreakerReadyCounts();
  const ready = outputReady && counts.total > 0 && counts.ready === counts.total;

  if (runBtn) {
    runBtn.disabled = circuitBreakerState.running || !ready;
  }

  if (addPanelBtn) addPanelBtn.disabled = circuitBreakerState.running;
  if (browseNewBtn) browseNewBtn.disabled = circuitBreakerState.running;
  if (browseExistingBtn) browseExistingBtn.disabled = circuitBreakerState.running;
  if (clearNewBtn) {
    clearNewBtn.disabled =
      circuitBreakerState.running || !circuitBreakerState.newOutputPath;
  }
  if (clearExistingBtn) {
    clearExistingBtn.disabled =
      circuitBreakerState.running || !circuitBreakerState.existingOutputPath;
  }

  if (breakerDrop) {
    breakerDrop.dataset.hasFile =
      activePanel && (activePanel.breakerPath || activePanel.breakerFile)
        ? "true"
        : "false";
    breakerDrop.classList.toggle("is-disabled", circuitBreakerState.running);
  }
  if (directoryDrop) {
    directoryDrop.dataset.hasFile =
      activePanel && (activePanel.directoryPath || activePanel.directoryFile)
        ? "true"
        : "false";
    directoryDrop.classList.toggle("is-disabled", circuitBreakerState.running);
  }

  if (modeNew) modeNew.disabled = circuitBreakerState.running;
  if (modeExisting) modeExisting.disabled = circuitBreakerState.running;
  if (panelNameInput) panelNameInput.disabled = circuitBreakerState.running || !activePanel;

  if (runningOverlay) {
    runningOverlay.hidden = !circuitBreakerState.running;
  }

  if (circuitBreakerState.running) {
    setCircuitBreakerStatus(
      `Running ${counts.total} panel${counts.total === 1 ? "" : "s"} in the background...`
    );
  } else if (!outputReady) {
    setCircuitBreakerStatus("Choose where to save or select an existing schedule file.");
  } else if (ready) {
    setCircuitBreakerStatus(
      `Ready to run ${counts.total} panel${counts.total === 1 ? "" : "s"}.`
    );
  } else {
    setCircuitBreakerStatus(
      `${counts.ready}/${counts.total} panel${counts.total === 1 ? "" : "s"} ready. Each panel needs breaker and directory photos.`
    );
  }
}

function resetCircuitBreakerForm() {
  circuitBreakerState.panels = [];
  circuitBreakerState.activePanelId = "";
  circuitBreakerState.nextPanelNumber = 1;
  circuitBreakerState.outputMode = "new";
  circuitBreakerState.newOutputExtension = "xlsx";
  circuitBreakerState.newOutputPath = "";
  circuitBreakerState.existingOutputPath = "";
  circuitBreakerState.running = false;
  ensureCircuitBreakerPanels();
  updateCircuitBreakerUi();
}

function setCircuitBreakerFile(kind, file) {
  const panel = getActiveCircuitBreakerPanel();
  if (!panel) return;
  if (kind === "breaker") {
    panel.breakerFile = file || null;
    if (file) panel.breakerPath = "";
  } else {
    panel.directoryFile = file || null;
    if (file) panel.directoryPath = "";
  }
  updateCircuitBreakerUi();
}

async function selectCircuitBreakerImage(kind) {
  const panel = getActiveCircuitBreakerPanel();
  if (!panel) return;
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
        panel.breakerPath = result.paths[0];
        panel.breakerFile = null;
      } else {
        panel.directoryPath = result.paths[0];
        panel.directoryFile = null;
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
      const outputExtension = normalizeCircuitBreakerOutputExtension(
        circuitBreakerState.newOutputExtension
      );
      const selection = await window.pywebview.api.select_template_save_location(
        null,
        "Panel_Schedule",
        outputExtension
      );
      if (selection?.status === "success" && selection.path) {
        circuitBreakerState.newOutputPath = selection.path;
        const selectedExtension = getCircuitBreakerOutputExtensionFromPath(
          selection.path
        );
        if (selectedExtension) {
          circuitBreakerState.newOutputExtension = selectedExtension;
        }
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
      file_types: ["Excel Files (*.xlsx;*.xls)"],
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
  ensureCircuitBreakerPanels();
  const outputPath = getCircuitBreakerOutputPath();
  if (!outputPath) {
    toast("Choose where to save or select an existing panel schedule.");
    return;
  }
  const pendingPanels = circuitBreakerState.panels.filter(
    (panel) => !isCircuitBreakerPanelReady(panel)
  );
  if (pendingPanels.length) {
    if (pendingPanels.length === 1) {
      toast(
        `Complete breaker and directory photos for ${getCircuitBreakerPanelDisplayName(
          pendingPanels[0]
        )}.`
      );
    } else {
      toast(
        `Complete breaker and directory photos for ${pendingPanels.length} panels before running.`
      );
    }
    return;
  }
  if (!window.pywebview?.api?.run_panel_schedule_background) {
    toast("Panel Schedule AI is unavailable in this environment.");
    return;
  }

  circuitBreakerState.running = true;
  updateCircuitBreakerUi();

  const panels = [];
  for (const panel of circuitBreakerState.panels) {
    panels.push({
      panelId: panel.id,
      panelName: panel.panelName?.trim() || panel.label || "PANEL",
      breakerPath: panel.breakerPath || "",
      directoryPath: panel.directoryPath || "",
      breakerUploads: panel.breakerFile
        ? [await fileToUploadPayload(panel.breakerFile)]
        : [],
      directoryUploads: panel.directoryFile
        ? [await fileToUploadPayload(panel.directoryFile)]
        : [],
    });
  }

  const firstPanel = panels[0] || {};
  const payload = {
    outputMode: circuitBreakerState.outputMode,
    outputPath,
    outputExtension:
      circuitBreakerState.outputMode === "new"
        ? normalizeCircuitBreakerOutputExtension(
          circuitBreakerState.newOutputExtension
        )
        : "",
    panels,
    breakerPath: firstPanel.breakerPath || "",
    directoryPath: firstPanel.directoryPath || "",
    breakerUploads: firstPanel.breakerUploads || [],
    directoryUploads: firstPanel.directoryUploads || [],
    panelName: firstPanel.panelName || "",
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
    toast(
      `Panel Schedule AI is processing ${panels.length} panel${panels.length === 1 ? "" : "s"
      } in the background.`
    );
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

  const parsedSuccess = Number(payload.successCount);
  const parsedFailure = Number(payload.failureCount);
  const successCount = Number.isFinite(parsedSuccess)
    ? Math.max(0, parsedSuccess)
    : payload.status === "success"
      ? 1
      : 0;
  const failureCount = Number.isFinite(parsedFailure)
    ? Math.max(0, parsedFailure)
    : payload.status === "success"
      ? 0
      : 1;
  const results = Array.isArray(payload.results) ? payload.results : [];

  if (payload.status === "success") {
    let message = payload.message;
    if (!message) {
      if (failureCount > 0) {
        message = `Panel Schedule AI finished: ${successCount} succeeded, ${failureCount} failed.`;
      } else {
        message = `Panel Schedule AI complete: ${successCount} panel${successCount === 1 ? "" : "s"} added.`;
      }
    } else if (failureCount > 0) {
      const failedNames = results
        .filter((item) => item?.status === "error")
        .map((item) => item?.panelName || item?.panelId)
        .filter(Boolean);
      if (failedNames.length) {
        const listed = failedNames.slice(0, 2).join(", ");
        message = `${message} Failed: ${listed}${failedNames.length > 2 ? ", ..." : ""}.`;
      }
    }
    toast(message, failureCount > 0 ? 6000 : 3500);

    const folder =
      payload.outputFolder ||
      (payload.outputPath ? payload.outputPath.split(/[\\/]/).slice(0, -1).join("\\") : "");
    if (successCount > 0 && folder && window.pywebview?.api?.open_path) {
      try {
        await window.pywebview.api.open_path(folder);
      } catch (e) {
        // Ignore open errors.
      }
    }
  } else {
    const baseMessage = payload.message || "Panel Schedule AI failed.";
    const failedNames = results
      .filter((item) => item?.status === "error")
      .map((item) => item?.panelName || item?.panelId)
      .filter(Boolean);
    const listed = failedNames.length
      ? ` Failed: ${failedNames.slice(0, 2).join(", ")}${failedNames.length > 2 ? ", ..." : ""}.`
      : "";
    toast(`${baseMessage}${listed}`, 6000);
  }
};

const debouncedSaveLightingSchedule = debounce(() => save(), 400);
const debouncedSaveTitle24 = debounce(() => save(), 400);

function autoResizeTextarea(textarea) {
  if (!textarea) return;
  textarea.style.height = "auto";
  textarea.style.height = `${textarea.scrollHeight}px`;
}

function normalizeTitle24Text(value) {
  return String(value ?? "").trim();
}

function normalizeTitle24Stories(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return null;
  const parsed = Number(raw);
  if (!Number.isInteger(parsed) || parsed <= 0) return null;
  return parsed;
}

function normalizeTitle24SquareFeet(value) {
  if (value === "" || value == null) return 0;
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed < 0) return 0;
  return Number(parsed.toFixed(4));
}

function createTitle24RoomAreaRow(seed = {}) {
  return {
    roomType: normalizeTitle24Text(seed.roomType ?? seed.RoomType),
    squareFeet: normalizeTitle24SquareFeet(seed.squareFeet ?? seed.SquareFeet),
  };
}

function computeTitle24RoomAreaTotal(rows = []) {
  const total = rows.reduce(
    (sum, row) => sum + normalizeTitle24SquareFeet(row?.squareFeet),
    0
  );
  return Number(total.toFixed(4));
}

function clampTitle24Page(page) {
  const parsed = Number(page);
  if (!Number.isInteger(parsed)) return 1;
  return Math.min(Math.max(parsed, 1), TITLE24_PAGE_COUNT);
}

function createDefaultTitle24() {
  return {
    schemaVersion: TITLE24_SCHEMA_VERSION,
    currentPage: 1,
    projectScope: {
      occupancyType: "",
      projectScopeType: "",
      lightingSystemType: "",
      aboveGradeStories: null,
    },
    roomAreas: {
      sourcePath: "",
      importedAtUtc: "",
      rows: [],
      totalSquareFeet: 0,
    },
  };
}

function normalizeTitle24(title24) {
  const normalized = title24 && typeof title24 === "object" ? title24 : {};
  const projectScope =
    normalized.projectScope && typeof normalized.projectScope === "object"
      ? normalized.projectScope
      : {};
  const roomAreas =
    normalized.roomAreas && typeof normalized.roomAreas === "object"
      ? normalized.roomAreas
      : {};
  const rows = Array.isArray(roomAreas.rows)
    ? roomAreas.rows.map((row) => createTitle24RoomAreaRow(row))
    : [];

  return {
    schemaVersion: TITLE24_SCHEMA_VERSION,
    currentPage: clampTitle24Page(normalized.currentPage),
    projectScope: {
      occupancyType: normalizeTitle24Text(projectScope.occupancyType),
      projectScopeType: normalizeTitle24Text(projectScope.projectScopeType),
      lightingSystemType: normalizeTitle24Text(projectScope.lightingSystemType),
      aboveGradeStories: normalizeTitle24Stories(projectScope.aboveGradeStories),
    },
    roomAreas: {
      sourcePath: normalizeTitle24Text(roomAreas.sourcePath),
      importedAtUtc: normalizeTitle24Text(roomAreas.importedAtUtc),
      rows,
      totalSquareFeet: computeTitle24RoomAreaTotal(rows),
    },
  };
}

function needsTitle24Migration(title24) {
  if (!title24 || typeof title24 !== "object") return true;
  if (title24.schemaVersion == null) return true;
  if (title24.currentPage == null) return true;
  if (!title24.projectScope || typeof title24.projectScope !== "object")
    return true;
  if (!title24.roomAreas || typeof title24.roomAreas !== "object") return true;
  if (!Array.isArray(title24.roomAreas.rows)) return true;
  if (title24.roomAreas.totalSquareFeet == null) return true;
  return TITLE24_SCOPE_OPTION_FIELDS.some(
    (field) => title24.projectScope[field] == null
  );
}

function ensureTitle24(project) {
  if (!project) return { title24: null, created: false, migrated: false };
  const existing = project.title24;
  const created = !existing;
  const migrated = needsTitle24Migration(existing);
  const title24 = normalizeTitle24(existing || createDefaultTitle24());
  project.title24 = title24;
  return { title24, created, migrated };
}

function createDefaultTitle24ScopeOptionsDataset() {
  return {
    schemaVersion: TITLE24_SCHEMA_VERSION,
    isPlaceholder: true,
    source: "EnergyCodeAce Indoor Lighting Compliance Form",
    capturedAtUtc: "",
    fields: {
      occupancyType: [],
      projectScopeType: [],
      lightingSystemType: [],
    },
  };
}

function normalizeTitle24ScopeOptionEntry(entry) {
  if (entry == null) return null;
  if (typeof entry === "string") {
    const value = normalizeTitle24Text(entry);
    if (!value) return null;
    return { value, label: value };
  }
  if (typeof entry !== "object") return null;
  const value = normalizeTitle24Text(entry.value);
  const label = normalizeTitle24Text(entry.label ?? entry.value);
  if (!value || !label) return null;
  return { value, label };
}

function normalizeTitle24ScopeOptionsDataset(rawDataset) {
  const defaults = createDefaultTitle24ScopeOptionsDataset();
  const raw =
    rawDataset && typeof rawDataset === "object" ? rawDataset : defaults;
  const fields = raw.fields && typeof raw.fields === "object" ? raw.fields : {};

  const normalizedFields = {};
  TITLE24_SCOPE_OPTION_FIELDS.forEach((field) => {
    const source = Array.isArray(fields[field]) ? fields[field] : [];
    const seen = new Set();
    normalizedFields[field] = source
      .map((entry) => normalizeTitle24ScopeOptionEntry(entry))
      .filter((entry) => {
        if (!entry) return false;
        const key = entry.value.toLowerCase();
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
  });

  return {
    schemaVersion:
      normalizeTitle24Text(raw.schemaVersion) || TITLE24_SCHEMA_VERSION,
    isPlaceholder: raw.isPlaceholder !== false,
    source: normalizeTitle24Text(raw.source) || defaults.source,
    capturedAtUtc: normalizeTitle24Text(raw.capturedAtUtc),
    fields: normalizedFields,
  };
}

function getMissingTitle24ScopeFields(dataset) {
  const fields =
    dataset?.fields && typeof dataset.fields === "object"
      ? dataset.fields
      : {};
  return TITLE24_SCOPE_OPTION_FIELDS.filter(
    (field) => !Array.isArray(fields[field]) || !fields[field].length
  );
}

function isTitle24ScopeOptionsBlocked(dataset) {
  if (!dataset) return true;
  if (dataset.isPlaceholder) return true;
  return getMissingTitle24ScopeFields(dataset).length > 0;
}

async function loadTitle24ScopeOptionsDataset({ force = false } = {}) {
  if (title24ScopeOptionsDataset && !force) {
    return title24ScopeOptionsDataset;
  }
  if (!force && title24ScopeOptionsLoadingPromise) {
    return title24ScopeOptionsLoadingPromise;
  }

  title24ScopeOptionsLoadingPromise = (async () => {
    const fallback = createDefaultTitle24ScopeOptionsDataset();
    try {
      const response = await fetch(TITLE24_SCOPE_OPTIONS_FILE_PATH, {
        cache: "no-store",
      });
      if (!response.ok) {
        throw new Error(`HTTP ${response.status} while loading options file.`);
      }
      const payload = await response.json();
      const normalized = normalizeTitle24ScopeOptionsDataset(payload);
      title24ScopeOptionsDataset = normalized;
      title24ScopeOptionsLoadError = "";
      return normalized;
    } catch (error) {
      console.warn("Failed to load Title 24 scope options dataset:", error);
      title24ScopeOptionsDataset = normalizeTitle24ScopeOptionsDataset(fallback);
      title24ScopeOptionsLoadError =
        error?.message ||
        "Could not load options file from assets/title24 folder.";
      return title24ScopeOptionsDataset;
    } finally {
      title24ScopeOptionsLoadingPromise = null;
    }
  })();

  return title24ScopeOptionsLoadingPromise;
}

function normalizeLightingScheduleText(value) {
  return String(value ?? "").replace(/\r\n?/g, "\n");
}

function createDefaultLightingScheduleSyncMeta() {
  return {
    targetDwgPath: "",
    lastSyncFilePath: "",
    lastPulledAtUtc: "",
    lastPushedAtUtc: "",
    lastPulledFingerprint: "",
    lastPushedFingerprint: "",
    lastSyncSource: "",
  };
}

function createLightingScheduleRow(seed = {}) {
  const row = {};
  LIGHTING_SCHEDULE_FIELDS.forEach((field) => {
    row[field] = normalizeLightingScheduleText(seed[field]);
  });
  return row;
}

function createDefaultLightingSchedule() {
  return {
    rows: [createLightingScheduleRow()],
    generalNotes: LIGHTING_SCHEDULE_DEFAULT_GENERAL_NOTES,
    notes: LIGHTING_SCHEDULE_DEFAULT_NOTES,
    ...createDefaultLightingScheduleSyncMeta(),
  };
}

function normalizeLightingSchedule(schedule) {
  const normalized = schedule && typeof schedule === "object" ? schedule : {};
  const syncDefaults = createDefaultLightingScheduleSyncMeta();
  if (!Array.isArray(normalized.rows) || normalized.rows.length === 0) {
    normalized.rows = [createLightingScheduleRow()];
  } else {
    normalized.rows = normalized.rows.map((row) => createLightingScheduleRow(row));
  }
  if (normalized.generalNotes == null) {
    normalized.generalNotes = LIGHTING_SCHEDULE_DEFAULT_GENERAL_NOTES;
  } else {
    normalized.generalNotes = normalizeLightingScheduleText(
      normalized.generalNotes
    );
  }
  if (normalized.notes == null) {
    normalized.notes = LIGHTING_SCHEDULE_DEFAULT_NOTES;
  } else {
    normalized.notes = normalizeLightingScheduleText(normalized.notes);
  }
  Object.keys(syncDefaults).forEach((key) => {
    normalized[key] = normalizeLightingScheduleText(normalized[key] ?? "");
  });
  return normalized;
}

function buildCanonicalLightingSchedule(schedule) {
  const normalized = normalizeLightingSchedule(schedule || {});
  return {
    rows: normalized.rows.map((row) => createLightingScheduleRow(row)),
    generalNotes: normalizeLightingScheduleText(normalized.generalNotes),
    notes: normalizeLightingScheduleText(normalized.notes),
  };
}

function stableStringify(value) {
  if (value === null || value === undefined) {
    return "null";
  }
  if (typeof value !== "object") {
    return JSON.stringify(value);
  }
  if (Array.isArray(value)) {
    return `[${value.map((item) => stableStringify(item)).join(",")}]`;
  }
  const keys = Object.keys(value).sort();
  const parts = keys.map(
    (key) => `${JSON.stringify(key)}:${stableStringify(value[key])}`
  );
  return `{${parts.join(",")}}`;
}

function hashFNV1a(text) {
  let hash = 0x811c9dc5;
  for (let i = 0; i < text.length; i += 1) {
    hash ^= text.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193) >>> 0;
  }
  return hash.toString(16).padStart(8, "0");
}

function computeLightingScheduleFingerprint(schedule) {
  const canonical = buildCanonicalLightingSchedule(schedule);
  const serialized = stableStringify(canonical);
  return `fnv1a:${hashFNV1a(serialized)}`;
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

function needsLightingScheduleMigration(schedule) {
  if (!schedule || typeof schedule !== "object") return true;
  if (!Array.isArray(schedule.rows) || schedule.rows.length === 0) return true;
  if (schedule.generalNotes == null || schedule.notes == null) return true;
  const syncDefaults = createDefaultLightingScheduleSyncMeta();
  return Object.keys(syncDefaults).some((key) => schedule[key] == null);
}

function getActiveLightingScheduleProject() {
  if (lightingScheduleProjectIndex == null) return null;
  return db[lightingScheduleProjectIndex] || null;
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
    lightingScheduleSyncStatusMessage = "";
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

function normalizeLightingSchedulePath(rawPath) {
  if (!rawPath) return "";
  let path = String(rawPath).trim().replace(/^['"]+|['"]+$/g, "");
  if (!path) return "";
  path = path.replace(/\//g, "\\").replace(/\\+$/g, "");
  return path;
}

function getLightingSchedulePathKey(rawPath) {
  return normalizeLightingSchedulePath(rawPath).toLowerCase();
}

function getFileNameFromPath(rawPath) {
  const path = normalizeLightingSchedulePath(rawPath);
  if (!path) return "";
  const idx = path.lastIndexOf("\\");
  if (idx < 0) return path;
  return path.slice(idx + 1);
}

function getFolderFromPath(rawPath) {
  const path = normalizeLightingSchedulePath(rawPath);
  if (!path) return "";
  const idx = path.lastIndexOf("\\");
  if (idx <= 0) return "";
  return path.slice(0, idx);
}

function getLightingScheduleSyncFilePathFromDwg(dwgPath) {
  const folder = getFolderFromPath(dwgPath);
  if (!folder) return "";
  return `${folder}\\${LIGHTING_SCHEDULE_SYNC_FILE_NAME}`;
}

function formatLightingScheduleSyncTimestamp(rawValue) {
  const value = String(rawValue || "").trim();
  if (!value) return "";
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return value;
  return parsed.toLocaleString();
}

function buildLightingScheduleSyncDefaultMessage(project, schedule) {
  if (!project || !schedule) {
    return "Select a project and target DWG to enable sync.";
  }
  const targetDwgPath = normalizeLightingSchedulePath(schedule.targetDwgPath);
  if (!targetDwgPath) {
    return "Choose Target DWG. Pull flow: run LFSPULL in AutoCAD and select table, then click Pull. Push flow: click Push, then run LFSPUSH in AutoCAD and select table.";
  }

  const syncPath = getLightingScheduleSyncFilePathFromDwg(targetDwgPath);
  const segments = [];
  if (syncPath) segments.push(`Sync file: ${syncPath}`);
  if (schedule.lastPulledAtUtc) {
    segments.push(`Last pull: ${formatLightingScheduleSyncTimestamp(schedule.lastPulledAtUtc)}`);
  }
  if (schedule.lastPushedAtUtc) {
    segments.push(`Last push: ${formatLightingScheduleSyncTimestamp(schedule.lastPushedAtUtc)}`);
  }
  segments.push("Pull: run LFSPULL in AutoCAD, then click Pull.");
  segments.push("Push: click Push, then run LFSPUSH in AutoCAD.");
  return segments.join(" ");
}

function setLightingScheduleSyncStatus(message) {
  lightingScheduleSyncStatusMessage = String(message || "").trim();
  const project = getActiveLightingScheduleProject();
  const schedule = getActiveLightingSchedule();
  renderLightingScheduleSyncControls(project, schedule);
}

function renderLightingScheduleSyncControls(project, schedule) {
  const targetInput = document.getElementById("lightingScheduleTargetDwg");
  const browseBtn = document.getElementById("lightingScheduleBrowseDwg");
  const clearBtn = document.getElementById("lightingScheduleClearDwg");
  const pullBtn = document.getElementById("lightingSchedulePullBtn");
  const pushBtn = document.getElementById("lightingSchedulePushBtn");
  const statusEl = document.getElementById("lightingScheduleSyncStatus");
  if (!targetInput && !statusEl) return;

  const hasProject = !!project && !!schedule;
  const targetPath = hasProject
    ? normalizeLightingSchedulePath(schedule.targetDwgPath)
    : "";

  if (targetInput) {
    targetInput.value = targetPath;
    targetInput.title = targetPath || "";
    targetInput.disabled = !hasProject;
  }
  if (browseBtn) browseBtn.disabled = !hasProject;
  if (clearBtn) clearBtn.disabled = !hasProject || !targetPath;
  if (pullBtn) pullBtn.disabled = !hasProject || !targetPath;
  if (pushBtn) pushBtn.disabled = !hasProject || !targetPath;

  if (statusEl) {
    statusEl.textContent =
      lightingScheduleSyncStatusMessage ||
      buildLightingScheduleSyncDefaultMessage(project, schedule);
  }
}

function updateLightingScheduleTargetDwgPath(project, schedule, dwgPath) {
  if (!project || !schedule) return;
  const normalizedPath = normalizeLightingSchedulePath(dwgPath);
  schedule.targetDwgPath = normalizedPath;
  schedule.lastSyncFilePath = normalizedPath
    ? getLightingScheduleSyncFilePathFromDwg(normalizedPath)
    : "";
}

function getLightingScheduleProjectMatchResult(project, syncProject, targetDwgPath) {
  const desktopProjectId = String(project?.id || "").trim();
  const syncProjectId = String(syncProject?.projectId || "").trim();
  if (desktopProjectId && syncProjectId) {
    return {
      matched: desktopProjectId === syncProjectId,
      mode: "projectId",
      desktopValue: desktopProjectId,
      syncValue: syncProjectId,
    };
  }

  const desktopBaseKey = getProjectBaseKey(project?.path || targetDwgPath || "");
  const syncBaseKey = getProjectBaseKey(
    syncProject?.projectBasePath || syncProject?.dwgPath || ""
  );
  if (desktopBaseKey && syncBaseKey) {
    return {
      matched: desktopBaseKey === syncBaseKey,
      mode: "projectBasePath",
      desktopValue: getProjectBasePath(project?.path || targetDwgPath || ""),
      syncValue: getProjectBasePath(
        syncProject?.projectBasePath || syncProject?.dwgPath || ""
      ),
    };
  }

  return {
    matched: false,
    mode: "unknown",
    desktopValue: "",
    syncValue: "",
  };
}

function normalizeIncomingLightingSyncPayload(rawPayload) {
  const payload =
    rawPayload && typeof rawPayload === "object" ? rawPayload : null;
  if (!payload) {
    return null;
  }
  const rows =
    Array.isArray(payload?.schedule?.rows) && payload.schedule.rows.length
      ? payload.schedule.rows.map((row) => createLightingScheduleRow(row))
      : [createLightingScheduleRow()];
  const schedule = {
    rows,
    generalNotes: normalizeLightingScheduleText(payload?.schedule?.generalNotes),
    notes: normalizeLightingScheduleText(payload?.schedule?.notes),
  };
  return {
    schemaVersion: String(payload.schemaVersion || "").trim(),
    metadata:
      payload.metadata && typeof payload.metadata === "object"
        ? payload.metadata
        : {},
    project:
      payload.project && typeof payload.project === "object"
        ? payload.project
        : {},
    table:
      payload.table && typeof payload.table === "object" ? payload.table : {},
    schedule,
  };
}

function buildLightingScheduleSyncPayload(project, schedule) {
  const targetDwgPath = normalizeLightingSchedulePath(schedule.targetDwgPath);
  const canonicalSchedule = buildCanonicalLightingSchedule(schedule);
  const fingerprint = computeLightingScheduleFingerprint(canonicalSchedule);
  return {
    schemaVersion: LIGHTING_SCHEDULE_SYNC_SCHEMA_VERSION,
    metadata: {
      sourceApp: "desktop",
      generatedAtUtc: new Date().toISOString(),
      generatedBy: String(userSettings?.userName || "desktop").trim() || "desktop",
      fingerprint,
    },
    project: {
      projectId: String(project?.id || "").trim(),
      projectBasePath: getProjectBasePath(project?.path || targetDwgPath || ""),
      dwgPath: targetDwgPath,
      dwgName: getFileNameFromPath(targetDwgPath),
    },
    table: {
      tableHandle: "",
    },
    schedule: canonicalSchedule,
  };
}

async function browseLightingScheduleTargetDwg() {
  const project = getActiveLightingScheduleProject();
  const schedule = getActiveLightingSchedule();
  if (!project || !schedule) {
    toast("Select a project first.");
    return;
  }
  if (!window.pywebview?.api?.select_files) {
    toast("File picker is unavailable.");
    return;
  }

  try {
    const result = await window.pywebview.api.select_files({
      allow_multiple: false,
      file_types: ["Drawing Files (*.dwg)"],
    });
    if (result?.status !== "success" || !Array.isArray(result.paths) || !result.paths.length) {
      return;
    }
    updateLightingScheduleTargetDwgPath(project, schedule, result.paths[0]);
    lightingScheduleSyncStatusMessage = "";
    renderLightingScheduleSyncControls(project, schedule);
    save();
  } catch (error) {
    reportClientError("Failed to browse target DWG", error);
    toast("Unable to select DWG path.");
  }
}

function clearLightingScheduleTargetDwg() {
  const project = getActiveLightingScheduleProject();
  const schedule = getActiveLightingSchedule();
  if (!project || !schedule) return;
  updateLightingScheduleTargetDwgPath(project, schedule, "");
  lightingScheduleSyncStatusMessage = "";
  renderLightingScheduleSyncControls(project, schedule);
  save();
}

async function pullLightingScheduleFromAutoCAD() {
  const project = getActiveLightingScheduleProject();
  const schedule = getActiveLightingSchedule();
  if (!project || !schedule) {
    toast("Select a project first.");
    return;
  }
  const targetDwgPath = normalizeLightingSchedulePath(schedule.targetDwgPath);
  if (!targetDwgPath) {
    toast("Select Target DWG before pulling.");
    return;
  }
  if (!window.pywebview?.api?.get_lighting_schedule_sync) {
    toast("Desktop backend is missing lighting schedule sync APIs.");
    return;
  }

  const ready = confirm(
    "In AutoCAD, run LFSPULL and select the target lighting fixture table.\n\nPress OK after LFSPULL has completed to import into the desktop schedule."
  );
  if (!ready) {
    setLightingScheduleSyncStatus("Pull cancelled. Run LFSPULL first, then click Pull again.");
    return;
  }

  try {
    const response = await window.pywebview.api.get_lighting_schedule_sync(
      targetDwgPath
    );
    if (response?.status !== "success") {
      throw new Error(response?.message || "Unable to read sync file.");
    }
    if (!response.exists) {
      setLightingScheduleSyncStatus(
        `No sync file found at ${response.path}. Run LFSPULL in AutoCAD and then click Pull again.`
      );
      toast("Sync file not found. Run LFSPULL first.");
      return;
    }

    const payload = normalizeIncomingLightingSyncPayload(response.data);
    if (!payload) {
      throw new Error("Sync payload is empty or malformed.");
    }
    const incomingFingerprint =
      String(payload?.metadata?.fingerprint || "").trim() ||
      computeLightingScheduleFingerprint(payload.schedule);

    const match = getLightingScheduleProjectMatchResult(
      project,
      payload.project,
      targetDwgPath
    );
    if (!match.matched) {
      const keepGoing = confirm(
        `Project mismatch detected (${match.mode}).\nDesktop: ${match.desktopValue || "(unknown)"}\nSync file: ${match.syncValue || "(unknown)"}\n\nPress OK to continue importing anyway.`
      );
      if (!keepGoing) {
        setLightingScheduleSyncStatus("Pull cancelled because project match validation failed.");
        return;
      }
    }

    const currentFingerprint = computeLightingScheduleFingerprint(schedule);
    if (currentFingerprint !== incomingFingerprint) {
      const overwrite = confirm(
        "Desktop lighting schedule differs from AutoCAD sync data.\n\nPress OK to overwrite desktop schedule with pulled AutoCAD data."
      );
      if (!overwrite) {
        setLightingScheduleSyncStatus("Pull cancelled. Desktop schedule was not overwritten.");
        return;
      }
    }

    schedule.rows = payload.schedule.rows.map((row) => createLightingScheduleRow(row));
    schedule.generalNotes = payload.schedule.generalNotes;
    schedule.notes = payload.schedule.notes;
    schedule.lastSyncFilePath = String(response.path || "").trim();
    schedule.lastPulledAtUtc = new Date().toISOString();
    schedule.lastPulledFingerprint = incomingFingerprint;
    schedule.lastSyncSource = "autocad";

    lightingScheduleSyncStatusMessage = "";
    renderLightingSchedule(schedule);
    save();
    setLightingScheduleSyncStatus(
      `Pulled AutoCAD sync data from ${response.path}. Schedule is now updated in desktop for project ${project.id || project.name || "(unnamed)"}.`
    );
    toast("Lighting schedule pulled from AutoCAD sync file.");
  } catch (error) {
    reportClientError("Failed to pull lighting schedule sync", error);
    setLightingScheduleSyncStatus(
      `Pull failed: ${error?.message || "Unknown error"}.`
    );
    toast("Pull failed. See sync status for details.");
  }
}

async function pushLightingScheduleToAutoCAD() {
  const project = getActiveLightingScheduleProject();
  const schedule = getActiveLightingSchedule();
  if (!project || !schedule) {
    toast("Select a project first.");
    return;
  }
  const targetDwgPath = normalizeLightingSchedulePath(schedule.targetDwgPath);
  if (!targetDwgPath) {
    toast("Select Target DWG before pushing.");
    return;
  }
  if (!window.pywebview?.api?.save_lighting_schedule_sync) {
    toast("Desktop backend is missing lighting schedule sync APIs.");
    return;
  }

  try {
    const payload = buildLightingScheduleSyncPayload(project, schedule);
    const nextFingerprint = String(payload?.metadata?.fingerprint || "").trim();

    if (window.pywebview?.api?.get_lighting_schedule_sync) {
      const existing = await window.pywebview.api.get_lighting_schedule_sync(
        targetDwgPath
      );
      if (existing?.status === "success" && existing.exists) {
        const existingPayload = normalizeIncomingLightingSyncPayload(existing.data);
        const existingFingerprint =
          String(existingPayload?.metadata?.fingerprint || "").trim() ||
          computeLightingScheduleFingerprint(existingPayload?.schedule || {});
        if (existingFingerprint !== nextFingerprint) {
          const overwrite = confirm(
            "Existing AutoCAD sync file differs from desktop lighting schedule.\n\nPress OK to overwrite sync file with desktop data."
          );
          if (!overwrite) {
            setLightingScheduleSyncStatus("Push cancelled. Existing sync file was kept.");
            return;
          }
        }
      }
    }

    const response = await window.pywebview.api.save_lighting_schedule_sync(
      targetDwgPath,
      payload
    );
    if (response?.status !== "success") {
      throw new Error(response?.message || "Failed to write sync file.");
    }

    schedule.lastSyncFilePath = String(response.path || "").trim();
    schedule.lastPushedAtUtc = new Date().toISOString();
    schedule.lastPushedFingerprint = nextFingerprint;
    schedule.lastSyncSource = "desktop";
    lightingScheduleSyncStatusMessage = "";
    renderLightingScheduleSyncControls(project, schedule);
    save();
    setLightingScheduleSyncStatus(
      `Push complete. Run LFSPUSH in AutoCAD, then select the target table to apply ${response.path}.`
    );
    toast("Lighting schedule sync file written. Run LFSPUSH in AutoCAD.");
  } catch (error) {
    reportClientError("Failed to push lighting schedule sync", error);
    setLightingScheduleSyncStatus(
      `Push failed: ${error?.message || "Unknown error"}.`
    );
    toast("Push failed. See sync status for details.");
  }
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
    renderLightingScheduleSyncControls(null, null);
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
  renderLightingScheduleSyncControls(getActiveLightingScheduleProject(), schedule);
}

function setLightingScheduleProject(index) {
  if (!Number.isInteger(index) || !db[index]) {
    lightingScheduleProjectIndex = null;
    lightingScheduleSyncStatusMessage = "";
    renderLightingSchedule(null);
    return;
  }
  lightingScheduleProjectIndex = index;
  lightingScheduleSyncStatusMessage = "";
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
  const activeProject = getActiveLightingScheduleProject();
  if (activeProject) {
    renderLightingSchedule(ensureLightingSchedule(activeProject).schedule);
  }

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

function getActiveTitle24Project() {
  if (title24ProjectIndex == null) return null;
  return db[title24ProjectIndex] || null;
}

function getActiveTitle24() {
  if (title24ProjectIndex == null) return null;
  const project = db[title24ProjectIndex];
  if (!project) return null;
  return ensureTitle24(project).title24;
}

function getTitle24PageTitle(page) {
  const titles = {
    1: "Project Scope",
    2: "Lighting Systems",
    3: "Controls",
    4: "Review",
  };
  return titles[page] || `Page ${page}`;
}

function getTitle24ScopeFieldLabel(field) {
  const labels = {
    occupancyType: "Occupancy Type",
    projectScopeType: "Project Scope",
    lightingSystemType: "Lighting System Type",
  };
  return labels[field] || field;
}

function getTitle24ScopeBlockingMessage(dataset) {
  if (title24ScopeOptionsLoadError) {
    return `Strict dropdowns are locked because options could not be loaded (${title24ScopeOptionsLoadError}). Populate ${TITLE24_SCOPE_OPTIONS_FILE_PATH} with exact EnergyCodeAce options.`;
  }
  if (!dataset || dataset.isPlaceholder) {
    return `Strict dropdowns are locked until ${TITLE24_SCOPE_OPTIONS_FILE_PATH} is populated with exact EnergyCodeAce options.`;
  }
  const missing = getMissingTitle24ScopeFields(dataset);
  if (missing.length) {
    const labels = missing.map((field) => getTitle24ScopeFieldLabel(field));
    return `Strict dropdowns are locked because option lists are missing for: ${labels.join(", ")}.`;
  }
  return "";
}

function renderTitle24ScopeBlockingNotice(dataset) {
  const notice = document.getElementById("title24ScopeBlockedNotice");
  if (!notice) return;
  const blocked = isTitle24ScopeOptionsBlocked(dataset);
  if (!blocked) {
    notice.hidden = true;
    notice.textContent = "";
    return;
  }
  notice.hidden = false;
  notice.textContent = getTitle24ScopeBlockingMessage(dataset);
}

function populateTitle24ScopeSelect(
  select,
  options = [],
  selectedValue = "",
  disabled = false
) {
  if (!select) return;
  const selected = normalizeTitle24Text(selectedValue);
  const normalizedOptions = Array.isArray(options)
    ? options.map((option) => normalizeTitle24ScopeOptionEntry(option)).filter(Boolean)
    : [];

  select.innerHTML = "";
  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = disabled
    ? "Options locked until dataset is populated"
    : "Select an option";
  select.appendChild(placeholder);

  let hasSelected = false;
  normalizedOptions.forEach((option) => {
    const optEl = document.createElement("option");
    optEl.value = option.value;
    optEl.textContent = option.label;
    if (option.value === selected) hasSelected = true;
    select.appendChild(optEl);
  });

  if (selected && !hasSelected) {
    const savedOpt = document.createElement("option");
    savedOpt.value = selected;
    savedOpt.textContent = `${selected} (Saved)`;
    select.appendChild(savedOpt);
    hasSelected = true;
  }

  select.value = hasSelected ? selected : "";
  select.disabled = !!disabled;
}

function formatTitle24Timestamp(rawValue) {
  const value = normalizeTitle24Text(rawValue);
  if (!value) return "";
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return value;
  return parsed.toLocaleString();
}

function formatTitle24TotalSquareFeet(value) {
  const n = normalizeTitle24SquareFeet(value);
  return `${n.toLocaleString(undefined, { maximumFractionDigits: 2 })} sq ft`;
}

function updateTitle24RoomAreasTotal(title24) {
  if (!title24) return;
  title24.roomAreas.totalSquareFeet = computeTitle24RoomAreaTotal(
    title24.roomAreas.rows
  );
}

function renderTitle24RoomAreaSummary(roomAreas) {
  const sourcePathEl = document.getElementById("title24RoomAreasSourcePath");
  const importedAtEl = document.getElementById("title24RoomAreasImportedAt");
  const totalEl = document.getElementById("title24RoomAreasTotalSqFt");
  if (sourcePathEl) {
    const sourcePath = normalizeTitle24Text(roomAreas?.sourcePath);
    sourcePathEl.textContent = sourcePath || "(not imported)";
    sourcePathEl.title = sourcePath;
  }
  if (importedAtEl) {
    const importedAt = formatTitle24Timestamp(roomAreas?.importedAtUtc);
    importedAtEl.textContent = importedAt || "(not imported)";
  }
  if (totalEl) {
    totalEl.textContent = formatTitle24TotalSquareFeet(
      roomAreas?.totalSquareFeet || 0
    );
  }
}

function renderTitle24RoomAreaRows(title24) {
  const body = document.getElementById("title24RoomAreasRows");
  if (!body) return;
  body.innerHTML = "";

  const rows = Array.isArray(title24?.roomAreas?.rows)
    ? title24.roomAreas.rows
    : [];

  if (!rows.length) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 3;
    td.className = "title24-room-empty-row tiny muted";
    td.textContent =
      "No room areas imported yet. Click Import T24Output.json to load room types.";
    tr.appendChild(td);
    body.appendChild(tr);
    return;
  }

  rows.forEach((row, rowIndex) => {
    const tr = document.createElement("tr");

    const roomTypeTd = document.createElement("td");
    const roomTypeInput = document.createElement("input");
    roomTypeInput.type = "text";
    roomTypeInput.value = row.roomType || "";
    roomTypeInput.placeholder = "Room type";
    roomTypeInput.addEventListener("input", (e) => {
      row.roomType = String(e.target.value || "");
      debouncedSaveTitle24();
    });
    roomTypeTd.appendChild(roomTypeInput);

    const squareFeetTd = document.createElement("td");
    const squareFeetInput = document.createElement("input");
    squareFeetInput.type = "number";
    squareFeetInput.min = "0";
    squareFeetInput.step = "0.01";
    squareFeetInput.value =
      row.squareFeet > 0 ? String(normalizeTitle24SquareFeet(row.squareFeet)) : "";
    squareFeetInput.placeholder = "0";
    squareFeetInput.addEventListener("input", (e) => {
      const raw = String(e.target.value || "").trim();
      if (!raw) {
        row.squareFeet = 0;
        squareFeetInput.classList.remove("input-error");
      } else {
        const parsed = Number(raw);
        if (!Number.isFinite(parsed) || parsed < 0) {
          squareFeetInput.classList.add("input-error");
          return;
        }
        row.squareFeet = Number(parsed);
        squareFeetInput.classList.remove("input-error");
      }
      updateTitle24RoomAreasTotal(title24);
      renderTitle24RoomAreaSummary(title24.roomAreas);
      debouncedSaveTitle24();
    });
    squareFeetTd.appendChild(squareFeetInput);

    const actionsTd = document.createElement("td");
    actionsTd.className = "title24-room-actions-cell";
    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.className = "title24-room-remove";
    removeBtn.textContent = "X";
    removeBtn.title = "Remove row";
    removeBtn.setAttribute("aria-label", "Remove row");
    removeBtn.addEventListener("click", () => {
      removeTitle24RoomAreaRow(rowIndex);
    });
    actionsTd.appendChild(removeBtn);

    tr.append(roomTypeTd, squareFeetTd, actionsTd);
    body.appendChild(tr);
  });
}

function renderTitle24PageShell(title24) {
  const titleEl = document.getElementById("title24PageTitle");
  const indicatorEl = document.getElementById("title24PageIndicator");
  const prevBtn = document.getElementById("title24PrevPageBtn");
  const nextBtn = document.getElementById("title24NextPageBtn");

  const hasData = !!title24;
  const currentPage = hasData ? clampTitle24Page(title24.currentPage) : 1;
  if (hasData) title24.currentPage = currentPage;

  if (titleEl) titleEl.textContent = getTitle24PageTitle(currentPage);
  if (indicatorEl) {
    indicatorEl.textContent = `Page ${currentPage} of ${TITLE24_PAGE_COUNT}`;
  }
  if (prevBtn) prevBtn.disabled = !hasData || currentPage <= 1;
  if (nextBtn) nextBtn.disabled = !hasData || currentPage >= TITLE24_PAGE_COUNT;

  for (let page = 1; page <= TITLE24_PAGE_COUNT; page += 1) {
    const section = document.getElementById(`title24Page${page}`);
    if (!section) continue;
    section.hidden = !hasData || page !== currentPage;
  }
}

function renderTitle24ProjectScopePage(title24) {
  const dataset =
    title24ScopeOptionsDataset || createDefaultTitle24ScopeOptionsDataset();
  const blocked = isTitle24ScopeOptionsBlocked(dataset);

  renderTitle24ScopeBlockingNotice(dataset);

  populateTitle24ScopeSelect(
    document.getElementById("title24OccupancyTypeSelect"),
    dataset.fields?.occupancyType,
    title24?.projectScope?.occupancyType || "",
    blocked
  );
  populateTitle24ScopeSelect(
    document.getElementById("title24ProjectScopeTypeSelect"),
    dataset.fields?.projectScopeType,
    title24?.projectScope?.projectScopeType || "",
    blocked
  );
  populateTitle24ScopeSelect(
    document.getElementById("title24LightingSystemTypeSelect"),
    dataset.fields?.lightingSystemType,
    title24?.projectScope?.lightingSystemType || "",
    blocked
  );

  const storiesInput = document.getElementById("title24AboveGradeStoriesInput");
  if (storiesInput) {
    storiesInput.disabled = !title24;
    storiesInput.value =
      title24?.projectScope?.aboveGradeStories == null
        ? ""
        : String(title24.projectScope.aboveGradeStories);
    storiesInput.classList.remove("input-error");
  }
}

function renderTitle24Compliance(
  title24,
  emptyMessage = "Create a project to start Title 24 compliance data entry."
) {
  const emptyState = document.getElementById("title24ComplianceEmptyState");
  const content = document.getElementById("title24ComplianceContent");
  const importBtn = document.getElementById("title24ImportRoomAreasBtn");
  const addRowBtn = document.getElementById("title24RoomAreasAddRow");

  if (!title24) {
    if (emptyState) {
      emptyState.textContent = emptyMessage;
      emptyState.hidden = false;
    }
    if (content) content.hidden = true;
    if (importBtn) importBtn.disabled = true;
    if (addRowBtn) addRowBtn.disabled = true;
    renderTitle24PageShell(null);
    renderTitle24ScopeBlockingNotice(
      title24ScopeOptionsDataset || createDefaultTitle24ScopeOptionsDataset()
    );
    renderTitle24RoomAreaSummary({
      sourcePath: "",
      importedAtUtc: "",
      totalSquareFeet: 0,
    });
    const rowsBody = document.getElementById("title24RoomAreasRows");
    if (rowsBody) rowsBody.innerHTML = "";
    return;
  }

  if (emptyState) emptyState.hidden = true;
  if (content) content.hidden = false;
  if (importBtn) importBtn.disabled = false;
  if (addRowBtn) addRowBtn.disabled = false;

  renderTitle24PageShell(title24);
  renderTitle24ProjectScopePage(title24);
  renderTitle24RoomAreaRows(title24);
  renderTitle24RoomAreaSummary(title24.roomAreas);
}

function renderTitle24ProjectOptions(filterText = "") {
  const select = document.getElementById("title24ProjectSelect");
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
    title24ProjectIndex = null;
    renderTitle24Compliance(
      null,
      db.length
        ? "No projects match that search."
        : "Create a project to start Title 24 compliance data entry."
    );
    return;
  }

  const availableIndexes = matches.map(({ index }) => index);
  if (availableIndexes.includes(title24ProjectIndex)) {
    select.value = String(title24ProjectIndex);
    return;
  }

  const nextIndex = availableIndexes[0];
  setTitle24Project(nextIndex);
}

function setTitle24Project(index) {
  if (!Number.isInteger(index) || !db[index]) {
    title24ProjectIndex = null;
    renderTitle24Compliance(null);
    return;
  }
  title24ProjectIndex = index;
  const select = document.getElementById("title24ProjectSelect");
  if (select) select.value = String(index);
  const { title24, created, migrated } = ensureTitle24(db[index]);
  if (created || migrated) save();
  renderTitle24Compliance(title24);
}

function handleTitle24ScopeSelectChange(field, value) {
  const title24 = getActiveTitle24();
  if (!title24 || !title24.projectScope) return;
  if (!TITLE24_SCOPE_OPTION_FIELDS.includes(field)) return;
  title24.projectScope[field] = normalizeTitle24Text(value);
  save();
}

function updateTitle24AboveGradeStories(rawValue, { showInvalidToast = false } = {}) {
  const title24 = getActiveTitle24();
  if (!title24) return false;
  const input = document.getElementById("title24AboveGradeStoriesInput");
  const raw = String(rawValue ?? "").trim();

  if (!raw) {
    title24.projectScope.aboveGradeStories = null;
    if (input) input.classList.remove("input-error");
    save();
    return true;
  }

  const parsed = normalizeTitle24Stories(raw);
  if (parsed == null) {
    if (input) input.classList.add("input-error");
    if (showInvalidToast) {
      toast("Above grade stories must be a positive whole number.");
    }
    return false;
  }

  title24.projectScope.aboveGradeStories = parsed;
  if (input) input.classList.remove("input-error");
  save();
  return true;
}

function addTitle24RoomAreaRow() {
  const title24 = getActiveTitle24();
  if (!title24) return;
  title24.roomAreas.rows.push(createTitle24RoomAreaRow());
  updateTitle24RoomAreasTotal(title24);
  renderTitle24RoomAreaRows(title24);
  renderTitle24RoomAreaSummary(title24.roomAreas);
  save();
}

function removeTitle24RoomAreaRow(rowIndex) {
  const title24 = getActiveTitle24();
  if (!title24) return;
  if (rowIndex < 0 || rowIndex >= title24.roomAreas.rows.length) return;
  title24.roomAreas.rows.splice(rowIndex, 1);
  updateTitle24RoomAreasTotal(title24);
  renderTitle24RoomAreaRows(title24);
  renderTitle24RoomAreaSummary(title24.roomAreas);
  save();
}

function changeTitle24Page(delta) {
  const title24 = getActiveTitle24();
  if (!title24) return;
  const nextPage = clampTitle24Page((title24.currentPage || 1) + Number(delta || 0));
  if (nextPage === title24.currentPage) return;
  title24.currentPage = nextPage;
  renderTitle24PageShell(title24);
  save();
}

async function importTitle24OutputJson() {
  const title24 = getActiveTitle24();
  if (!title24) {
    toast("Select a project first.");
    return;
  }
  if (!window.pywebview?.api?.select_files) {
    toast("File picker is unavailable.");
    return;
  }
  if (!window.pywebview?.api?.read_t24_output_json) {
    toast("Desktop backend is missing read_t24_output_json API.");
    return;
  }

  try {
    const selection = await window.pywebview.api.select_files({
      allow_multiple: false,
      file_types: ["JSON Files (*.json)"],
    });
    if (
      selection?.status !== "success" ||
      !Array.isArray(selection.paths) ||
      !selection.paths.length
    ) {
      return;
    }

    const response = await window.pywebview.api.read_t24_output_json(
      selection.paths[0]
    );
    if (response?.status !== "success") {
      throw new Error(response?.message || "Failed to read T24Output.json.");
    }

    const incomingRows = Array.isArray(response?.data?.rows)
      ? response.data.rows.map((row) => createTitle24RoomAreaRow(row))
      : [];

    title24.roomAreas.rows = incomingRows;
    title24.roomAreas.sourcePath = normalizeTitle24Text(
      response.path || selection.paths[0]
    );
    title24.roomAreas.importedAtUtc = new Date().toISOString();
    updateTitle24RoomAreasTotal(title24);

    renderTitle24Compliance(title24);
    save();
    toast(
      `Imported ${incomingRows.length} room type entr${incomingRows.length === 1 ? "y" : "ies"} from T24Output.json.`
    );
  } catch (error) {
    reportClientError("Failed to import T24Output.json", error);
    toast(`Import failed: ${error?.message || "Unknown error."}`);
  }
}

async function openTitle24Compliance() {
  const dlg = document.getElementById("title24ComplianceDlg");
  if (!dlg) return;
  await loadTitle24ScopeOptionsDataset({ force: true });
  const projectSearch = document.getElementById("title24ProjectSearch");
  if (projectSearch) projectSearch.value = title24ProjectQuery;
  renderTitle24ProjectOptions(title24ProjectQuery);
  const activeProject = getActiveTitle24Project();
  if (activeProject) {
    renderTitle24Compliance(ensureTitle24(activeProject).title24);
  }
  if (!dlg.open) dlg.showModal();
}

function closeTitle24Compliance() {
  const dlg = document.getElementById("title24ComplianceDlg");
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
      requestAnimationFrame(updateStickyOffsets);
    } else if (tab === "timesheets") {
      renderTimesheets();
    } else if (tab === "templates") {
      renderTemplates();
    }
    updateProjectsBackToTopVisibility();
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
  const pinUrgentDeliverablesBtn = document.getElementById(
    "pinUrgentDeliverablesBtn"
  );
  if (pinUrgentDeliverablesBtn) {
    pinUrgentDeliverablesBtn.textContent = "";
    pinUrgentDeliverablesBtn.appendChild(createIcon(PIN_ICON_PATH, 16));
    pinUrgentDeliverablesBtn.onclick = () => pinUrgentDeliverables();
  }

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

  const workroomToolsSettingsBtn = document.getElementById(
    "workroomToolsSettingsBtn"
  );
  if (workroomToolsSettingsBtn) {
    workroomToolsSettingsBtn.onclick = (e) => {
      e.stopPropagation();
      syncWorkroomCadRoutingInputs();
      const dlg = document.getElementById("workroomToolsSettingsDlg");
      if (dlg && !dlg.open) dlg.showModal();
    };
  }

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
  document.getElementById("checklistsHelpBtn").onclick = () =>
    openExternalUrl(HELP_LINKS.checklists);
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
  const editDlg = document.getElementById("editDlg");
  if (editDlg) {
    editDlg.addEventListener("close", () => {
      if (modalEmailSession.active) flushModalEmailSession(false);
    });
  }

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
      const launchContext = consumePendingCadLaunchContext();
      console.debug("Workroom CAD launch context (publish):", launchContext);
      if (e.currentTarget.classList.contains("running")) return;
      if (!userSettings.autocadPath) {
        await showAutocadSelectModal();
        return;
      }
      e.currentTarget.classList.add("running");
      window.updateToolStatus("toolPublishDwgs", "Initializing...");
      if (launchContext) {
        await window.pywebview.api.run_publish_script(launchContext);
      } else {
        await window.pywebview.api.run_publish_script();
      }
    });

  document
    .getElementById("toolFreezeLayers")
    .addEventListener("click", async (e) => {
      const launchContext = consumePendingCadLaunchContext();
      console.debug("Workroom CAD launch context (freeze):", launchContext);
      if (e.currentTarget.classList.contains("running")) return;
      if (!userSettings.autocadPath) {
        await showAutocadSelectModal();
        return;
      }
      e.currentTarget.classList.add("running");
      window.updateToolStatus("toolFreezeLayers", "Initializing...");
      if (launchContext) {
        await window.pywebview.api.run_freeze_layers_script(launchContext);
      } else {
        await window.pywebview.api.run_freeze_layers_script();
      }
    });

  document
    .getElementById("toolThawLayers")
    .addEventListener("click", async (e) => {
      const launchContext = consumePendingCadLaunchContext();
      console.debug("Workroom CAD launch context (thaw):", launchContext);
      if (e.currentTarget.classList.contains("running")) return;
      if (!userSettings.autocadPath) {
        await showAutocadSelectModal();
        return;
      }
      e.currentTarget.classList.add("running");
      window.updateToolStatus("toolThawLayers", "Initializing...");
      if (launchContext) {
        await window.pywebview.api.run_thaw_layers_script(launchContext);
      } else {
        await window.pywebview.api.run_thaw_layers_script();
      }
    });

  document
    .getElementById("toolCleanXrefs")
    .addEventListener("click", async (e) => {
      const launchContext = consumePendingCadLaunchContext();
      if (e.currentTarget.classList.contains("running")) return;
      if (!userSettings.autocadPath) {
        await showAutocadSelectModal();
        return;
      }
      e.currentTarget.classList.add("running");
      window.updateToolStatus("toolCleanXrefs", "Initializing...");
      if (launchContext) {
        await window.pywebview.api.run_clean_xrefs_script(launchContext);
      } else {
        await window.pywebview.api.run_clean_xrefs_script();
      }
    });

  const narrativeTemplateBtn = document.getElementById(
    "toolCreateNarrativeTemplate"
  );
  if (narrativeTemplateBtn) {
    const handler = async () => {
      const launchContext = consumePendingCadLaunchContext();
      console.debug("Workroom launch context (narrative template):", launchContext);
      await handleTemplateToolSave("narrative", "Narrative of Changes", {
        launchContext,
        toolId: "toolCreateNarrativeTemplate",
      });
    };
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
    const handler = async () => {
      const launchContext = consumePendingCadLaunchContext();
      console.debug("Workroom launch context (plan check template):", launchContext);
      await handleTemplateToolSave("planCheck", "Plan Check Comments", {
        launchContext,
        toolId: "toolCreatePlanCheckTemplate",
      });
    };
    planCheckTemplateBtn.addEventListener("click", handler);
    planCheckTemplateBtn.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        handler();
      }
    });
  }

  const copyProjectLocallyBtn = document.getElementById("toolCopyProjectLocally");
  if (copyProjectLocallyBtn) {
    const handler = async () => {
      if (copyProjectLocallyBtn.classList.contains("running")) return;
      copyProjectLocallyBtn.classList.add("running");
      window.updateToolStatus("toolCopyProjectLocally", "Select server project folder...");

      try {
        const selection = await window.pywebview.api.select_folder();
        if (selection?.status === "error") {
          throw new Error(selection.message || "Failed to choose a folder.");
        }
        if (!selection || selection.status === "cancelled" || !selection.path) {
          const statusEl = copyProjectLocallyBtn.querySelector(".tool-card-status");
          if (statusEl) statusEl.textContent = "";
          copyProjectLocallyBtn.classList.remove("running");
          return;
        }

        window.updateToolStatus("toolCopyProjectLocally", "Copying project...");
        const result = await window.pywebview.api.copy_project_locally(selection.path);
        if (result?.status !== "success") {
          throw new Error(result?.message || "Failed to copy project locally.");
        }

        window.updateToolStatus("toolCopyProjectLocally", "DONE");
        toast("Project copied locally.");

        const missingFolders = Array.isArray(result?.missingServerFolders)
          ? result.missingServerFolders
          : [];
        if (missingFolders.length) {
          toast(
            `Missing on server (created empty locally): ${missingFolders.join(", ")}`,
            7000
          );
        }

        if (result?.localProjectPath) {
          await window.pywebview.api.open_path(result.localProjectPath);
        }
      } catch (e) {
        const message = e?.message || "Failed to copy project locally.";
        window.updateToolStatus("toolCopyProjectLocally", `ERROR: ${message}`);
      }
    };

    copyProjectLocallyBtn.addEventListener("click", handler);
    copyProjectLocallyBtn.addEventListener("keydown", (e) => {
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
    cbOutputModeNew.addEventListener("click", () => {
      circuitBreakerState.outputMode = "new";
      updateCircuitBreakerUi();
    });
  }

  const cbOutputModeExisting = document.getElementById("cbOutputModeExisting");
  if (cbOutputModeExisting) {
    cbOutputModeExisting.addEventListener("click", () => {
      circuitBreakerState.outputMode = "existing";
      updateCircuitBreakerUi();
    });
  }

  const cbAddPanelTabBtn = document.getElementById("cbAddPanelTabBtn");
  if (cbAddPanelTabBtn) {
    cbAddPanelTabBtn.addEventListener("click", addCircuitBreakerPanel);
  }

  const cbNewScheduleFormat = document.getElementById("cbNewScheduleFormat");
  if (cbNewScheduleFormat) {
    cbNewScheduleFormat.addEventListener("change", (e) => {
      if (circuitBreakerState.running) return;
      const selected = normalizeCircuitBreakerOutputExtension(e.target.value);
      circuitBreakerState.newOutputExtension = selected;
      circuitBreakerState.newOutputPath = "";
      updateCircuitBreakerUi();
    });
  }

  const cbBrowseNewScheduleBtn = document.getElementById("cbBrowseNewScheduleBtn");
  if (cbBrowseNewScheduleBtn) {
    cbBrowseNewScheduleBtn.addEventListener("click", async () => {
      await selectCircuitBreakerSchedulePath("new");
    });
  }

  const cbClearNewScheduleBtn = document.getElementById("cbClearNewScheduleBtn");
  if (cbClearNewScheduleBtn) {
    cbClearNewScheduleBtn.addEventListener("click", () => {
      if (circuitBreakerState.running) return;
      circuitBreakerState.newOutputPath = "";
      updateCircuitBreakerUi();
    });
  }

  const cbBrowseExistingScheduleBtn = document.getElementById(
    "cbBrowseExistingScheduleBtn"
  );
  if (cbBrowseExistingScheduleBtn) {
    cbBrowseExistingScheduleBtn.addEventListener("click", async () => {
      await selectCircuitBreakerSchedulePath("existing");
    });
  }

  const cbClearExistingScheduleBtn = document.getElementById(
    "cbClearExistingScheduleBtn"
  );
  if (cbClearExistingScheduleBtn) {
    cbClearExistingScheduleBtn.addEventListener("click", () => {
      if (circuitBreakerState.running) return;
      circuitBreakerState.existingOutputPath = "";
      updateCircuitBreakerUi();
    });
  }

  const cbPanelNameInput = document.getElementById("cbPanelName");
  if (cbPanelNameInput) {
    cbPanelNameInput.addEventListener("input", (e) => {
      const panel = getActiveCircuitBreakerPanel();
      if (!panel) return;
      panel.panelName = e.target.value;
      updateCircuitBreakerUi();
    });
  }

  const cbRunPanelBtn = document.getElementById("cbRunPanelScheduleBtn");
  if (cbRunPanelBtn) {
    cbRunPanelBtn.addEventListener("click", runCircuitBreakerInBackground);
  }

  updateCircuitBreakerUi();

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

  const lightingScheduleBrowseDwgBtn = document.getElementById(
    "lightingScheduleBrowseDwg"
  );
  if (lightingScheduleBrowseDwgBtn) {
    lightingScheduleBrowseDwgBtn.addEventListener(
      "click",
      browseLightingScheduleTargetDwg
    );
  }

  const lightingScheduleClearDwgBtn = document.getElementById(
    "lightingScheduleClearDwg"
  );
  if (lightingScheduleClearDwgBtn) {
    lightingScheduleClearDwgBtn.addEventListener(
      "click",
      clearLightingScheduleTargetDwg
    );
  }

  const lightingSchedulePullBtn = document.getElementById(
    "lightingSchedulePullBtn"
  );
  if (lightingSchedulePullBtn) {
    lightingSchedulePullBtn.addEventListener(
      "click",
      pullLightingScheduleFromAutoCAD
    );
  }

  const lightingSchedulePushBtn = document.getElementById(
    "lightingSchedulePushBtn"
  );
  if (lightingSchedulePushBtn) {
    lightingSchedulePushBtn.addEventListener(
      "click",
      pushLightingScheduleToAutoCAD
    );
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

  const title24ComplianceBtn = document.getElementById("toolTitle24Compliance");
  if (title24ComplianceBtn) {
    title24ComplianceBtn.addEventListener("click", openTitle24Compliance);
    title24ComplianceBtn.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        openTitle24Compliance();
      }
    });
  }

  const title24ComplianceCloseBtn = document.getElementById(
    "title24ComplianceCloseBtn"
  );
  if (title24ComplianceCloseBtn) {
    title24ComplianceCloseBtn.addEventListener("click", closeTitle24Compliance);
  }

  const title24ComplianceDlg = document.getElementById("title24ComplianceDlg");
  if (title24ComplianceDlg) {
    title24ComplianceDlg.addEventListener("close", () => {
      save();
    });
  }

  const title24ProjectSelect = document.getElementById("title24ProjectSelect");
  if (title24ProjectSelect) {
    title24ProjectSelect.addEventListener("change", (e) => {
      const idx = Number(e.target.value);
      if (!Number.isNaN(idx)) setTitle24Project(idx);
    });
  }

  const title24ProjectSearch = document.getElementById("title24ProjectSearch");
  if (title24ProjectSearch) {
    title24ProjectSearch.addEventListener("input", (e) => {
      title24ProjectQuery = e.target.value;
      renderTitle24ProjectOptions(title24ProjectQuery);
    });
  }

  const title24PrevPageBtn = document.getElementById("title24PrevPageBtn");
  if (title24PrevPageBtn) {
    title24PrevPageBtn.addEventListener("click", () => changeTitle24Page(-1));
  }

  const title24NextPageBtn = document.getElementById("title24NextPageBtn");
  if (title24NextPageBtn) {
    title24NextPageBtn.addEventListener("click", () => changeTitle24Page(1));
  }

  const title24OccupancyTypeSelect = document.getElementById(
    "title24OccupancyTypeSelect"
  );
  if (title24OccupancyTypeSelect) {
    title24OccupancyTypeSelect.addEventListener("change", (e) => {
      handleTitle24ScopeSelectChange("occupancyType", e.target.value);
    });
  }

  const title24ProjectScopeTypeSelect = document.getElementById(
    "title24ProjectScopeTypeSelect"
  );
  if (title24ProjectScopeTypeSelect) {
    title24ProjectScopeTypeSelect.addEventListener("change", (e) => {
      handleTitle24ScopeSelectChange("projectScopeType", e.target.value);
    });
  }

  const title24LightingSystemTypeSelect = document.getElementById(
    "title24LightingSystemTypeSelect"
  );
  if (title24LightingSystemTypeSelect) {
    title24LightingSystemTypeSelect.addEventListener("change", (e) => {
      handleTitle24ScopeSelectChange("lightingSystemType", e.target.value);
    });
  }

  const title24AboveGradeStoriesInput = document.getElementById(
    "title24AboveGradeStoriesInput"
  );
  if (title24AboveGradeStoriesInput) {
    title24AboveGradeStoriesInput.addEventListener("input", (e) => {
      updateTitle24AboveGradeStories(e.target.value, { showInvalidToast: false });
    });
    title24AboveGradeStoriesInput.addEventListener("blur", (e) => {
      updateTitle24AboveGradeStories(e.target.value, { showInvalidToast: true });
    });
  }

  const title24ImportRoomAreasBtn = document.getElementById(
    "title24ImportRoomAreasBtn"
  );
  if (title24ImportRoomAreasBtn) {
    title24ImportRoomAreasBtn.addEventListener("click", importTitle24OutputJson);
  }

  const title24RoomAreasAddRowBtn = document.getElementById(
    "title24RoomAreasAddRow"
  );
  if (title24RoomAreasAddRowBtn) {
    title24RoomAreasAddRowBtn.addEventListener("click", addTitle24RoomAreaRow);
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
    const aiSpinner = document.getElementById("aiSpinner");
    const btnProcessEmail = document.getElementById("btnProcessEmail");
    const AI_EMAIL_TIMEOUT_MS = 120000;
    let timeoutId = null;
    aiSpinner.style.display = "flex";
    if (btnProcessEmail) btnProcessEmail.disabled = true;
    try {
      const aiRequest = window.pywebview.api.process_email_with_ai(
        txt,
        userSettings.apiKey,
        userSettings.userName,
        userSettings.discipline
      );
      const timeoutPromise = new Promise((_, reject) => {
        timeoutId = setTimeout(() => {
          reject(
            new Error(
              `AI request timed out after ${Math.round(
                AI_EMAIL_TIMEOUT_MS / 1000
              )} seconds. Please try again.`
            )
          );
        }, AI_EMAIL_TIMEOUT_MS);
      });
      const res = await Promise.race([aiRequest, timeoutPromise]);
      if (timeoutId) {
        clearTimeout(timeoutId);
        timeoutId = null;
      }
      if (res?.status === "success") {
        closeDlg("emailDlg");
        handleAiProjectResult(res.data || {});
      } else throw new Error(res?.message || "Failed to process email.");
    } catch (e) {
      toast("AI Error: " + (e?.message || "Unknown error."));
    } finally {
      if (timeoutId) clearTimeout(timeoutId);
      aiSpinner.style.display = "none";
      if (btnProcessEmail) btnProcessEmail.disabled = false;
    }
  };

  const aiNoMatchDlg = document.getElementById("aiNoMatchDlg");
  if (aiNoMatchDlg) {
    aiNoMatchDlg.addEventListener("close", () => {
      resetAiNoMatchState();
    });
  }

  const aiNoMatchCancelBtn = document.getElementById("aiNoMatchCancelBtn");
  if (aiNoMatchCancelBtn) {
    aiNoMatchCancelBtn.onclick = () => closeAiNoMatchDialog();
  }

  const aiNoMatchSearchBtn = document.getElementById("aiNoMatchSearchBtn");
  if (aiNoMatchSearchBtn) {
    aiNoMatchSearchBtn.onclick = () => openAiNoMatchSearch();
  }

  const aiNoMatchCreateNewBtn = document.getElementById("aiNoMatchCreateNewBtn");
  if (aiNoMatchCreateNewBtn) {
    aiNoMatchCreateNewBtn.onclick = () => {
      const rawAiData = aiNoMatchState.rawAiData || {};
      closeAiNoMatchDialog();
      openAiCreateNewProject(rawAiData);
    };
  }

  const aiNoMatchBackBtn = document.getElementById("aiNoMatchBackBtn");
  if (aiNoMatchBackBtn) {
    aiNoMatchBackBtn.onclick = () => setAiNoMatchDialogMode("choice");
  }

  const aiNoMatchSearchInput = document.getElementById("aiNoMatchSearchInput");
  if (aiNoMatchSearchInput) {
    aiNoMatchSearchInput.oninput = (e) => {
      setAiNoMatchSearchQuery(e.target.value);
    };
    aiNoMatchSearchInput.onkeydown = (e) => {
      if (e.key !== "Enter") return;
      e.preventDefault();
      if (
        aiNoMatchState.selectedProjectIndex < 0 &&
        aiNoMatchState.candidateProjectIndices.length === 1
      ) {
        selectAiNoMatchProject(aiNoMatchState.candidateProjectIndices[0]);
        return;
      }
      if (aiNoMatchState.selectedProjectIndex >= 0) {
        confirmAiNoMatchAddToProject();
      }
    };
  }

  const aiNoMatchCreateNewFallbackBtn = document.getElementById(
    "aiNoMatchCreateNewFallbackBtn"
  );
  if (aiNoMatchCreateNewFallbackBtn) {
    aiNoMatchCreateNewFallbackBtn.onclick = () => {
      const rawAiData = aiNoMatchState.rawAiData || {};
      closeAiNoMatchDialog();
      openAiCreateNewProject(rawAiData);
    };
  }

  const aiNoMatchAddBtn = document.getElementById("aiNoMatchAddBtn");
  if (aiNoMatchAddBtn) {
    aiNoMatchAddBtn.onclick = () => confirmAiNoMatchAddToProject();
  }

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

  ["settings_workroomAutoSelectCadFiles", "workroom_modal_autoSelectCadFiles"]
    .map((id) => document.getElementById(id))
    .filter(Boolean)
    .forEach((checkbox) => {
      checkbox.onchange = (e) => {
        userSettings.workroomAutoSelectCadFiles = e.target.checked;
        syncWorkroomCadRoutingInputs();
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
    initProjectsBackToTop();
    updateStickyOffsets();
    refreshAppUpdateStatus();
    await loadUserSettings();
    initThemeFromPreferences();

    if (isNewUser()) {
      hideMainApp();
      showOnboardingModal();
    } else {
      showMainApp();
      const [loadedDb, loadedNotes, loadedTimesheets, loadedTemplates, loadedChecklists] = await Promise.all([
        load(),
        loadNotes(),
        loadTimesheets(),
        loadTemplates(),
        loadChecklists(),
      ]);
      db = loadedDb;
      notesDb = loadedNotes || {};
      timesheetDb = loadedTimesheets || { weeks: {}, lastModified: null };
      templatesDb = loadedTemplates || { templates: [], defaultTemplatesInstalled: false, lastModified: null };
      checklistsDb = loadedChecklists || { checklists: [], lastModified: null };

      // Initialize checklist tab
      if (checklistsDb.checklists.length > 0) {
        activeChecklistTabId = checklistsDb.checklists[0].id;
      }
      renderNoteTabs();
      renderChecklistTabs();
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
