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
const FIREBASE_JS_SDK_VERSION = "12.7.0";
const FIREBASE_COMPAT_SCRIPT_URLS = [
  `https://www.gstatic.com/firebasejs/${FIREBASE_JS_SDK_VERSION}/firebase-app-compat.js`,
  `https://www.gstatic.com/firebasejs/${FIREBASE_JS_SDK_VERSION}/firebase-auth-compat.js`,
  `https://www.gstatic.com/firebasejs/${FIREBASE_JS_SDK_VERSION}/firebase-firestore-compat.js`,
];
const CLOUD_SYNC_TIMESHEETS_META_DOC_ID = "__meta__";
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
const PROJECT_ROOT_SEGMENT_REGEX =
  /^(\d{6})(?!\d)(?:\s*(?:[-_]\s*)?(.*))?$/;
const LIGHTING_PROJECT_SEGMENT_REGEX = PROJECT_ROOT_SEGMENT_REGEX;
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
const NOTE_ICON_PATH =
  "M14 2H6c-1.1 0-2 .9-2 2v16c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V8l-6-6zm1 7V3.5L20.5 9H15zM8 13h8v2H8v-2zm0 4h8v2H8v-2zm0-8h5v2H8V9z";
const PENCIL_ICON_PATH =
  "M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zm17.71-10.21c.39-.39.39-1.02 0-1.41l-2.34-2.34a1 1 0 0 0-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z";
const TRASH_ICON_PATH =
  "M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zm13-15h-3.5l-1-1h-5l-1 1H5v2h14V4z";
const MAIL_ICON_PATH =
  "M20 4H4c-1.1 0-2 .9-2 2v12a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V6c0-1.1-.9-2-2-2zm-.8 2L12 11.2 4.8 6h14.4zM4 18V7l7.4 5.2a1 1 0 0 0 1.2 0L20 7v11H4z";
// Checklist Icons
const CHECKLIST_ICON_PATH =
  "M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V5h14v14zM7 10h2v7H7zm4-3h2v10h-2zm4 6h2v4h-2z";
const CHECK_ICON_PATH =
  "M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z";
const MAX_DELIVERABLE_EMAIL_REFS = 3;

// Checklists Data
let checklistsDb = {
  checklists: [],
  templateOverrides: {},
  lastModified: null,
};

// Active checklist state for modal
let activeChecklistProject = null;
let activeChecklistDeliverable = null;
let activeChecklistTab = null;
let checklistDragState = null;
let checklistRowMenuState = null;
let checklistHandleMenuSuppressClickUntil = 0;
let activeWorkroomUtilityTab = "tasks";
let activeWorkroomPreDesignPage = 0;
let workroomSurveyDraftLoadPromise = null;
let workroomSurveyDraftProjectIndex = null;
let workroomDisciplineNoticeShown = false;
let workroomToolStatusState = {
  toolId: "",
  label: "",
  message: "Ready.",
  phase: "idle",
};
const ACTIVITY_STATUS = Object.freeze({
  RUNNING: "running",
  SUCCESS: "success",
  WARNING: "warning",
  ERROR: "error",
});
const activityTrayState = {
  items: [],
  collapsed: false,
  hasAutoExpanded: false,
  initialized: false,
};
const activeToolActivityIds = new Map();

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

const CHECKLIST_TEMPLATE_KEYS = {
  PRE_FLIGHT_ELECTRICAL: "pre_flight_electrical",
  ELECTRICAL_GENERAL: "electrical_general",
  ACIES_MECHANICAL: "acies_mechanical",
  ACIES_ELECTRICAL: "acies_electrical",
  ACIES_PLUMBING: "acies_plumbing",
};
const CHECKLIST_ROW_TYPES = {
  ITEM: "item",
  SUBHEADER: "subheader",
};
const PRE_FLIGHT_ELECTRICAL_CHECKLIST_NAME = "Pre-flight Electrical Checklist";
const ELECTRICAL_GENERAL_CHECKLIST_NAME = "Electrical General Checklist";
const ACIES_MECHANICAL_CHECKLIST_NAME = "Mechanical - ACIES Checklist";
const ACIES_ELECTRICAL_CHECKLIST_NAME = "Electrical - ACIES Checklist";
const ACIES_PLUMBING_CHECKLIST_NAME = "Plumbing - ACIES Checklist";
const LEGACY_DEFAULT_CHECKLIST_ID = "checklist_default";
const CHECKLIST_PROFILE_SCHEMA_VERSION = 1;
const ACIES_ELECTRICAL_SCOPE_PROFILE_KEY = CHECKLIST_TEMPLATE_KEYS.ACIES_ELECTRICAL;
const ACIES_ELECTRICAL_SCOPE_STATE_OPTIONS = [
  { value: "ca", label: "California" },
  { value: "wa", label: "Washington" },
  { value: "or", label: "Oregon" },
  { value: "nm", label: "New Mexico" },
  { value: "other", label: "Other / Multi-state" },
];
const ACIES_ELECTRICAL_SCOPE_CITY_AREA_OPTIONS = [
  { value: "san_francisco", label: "San Francisco" },
  { value: "san_jose", label: "San Jose" },
  { value: "santa_clara_county", label: "Santa Clara County" },
  { value: "milpitas", label: "Milpitas" },
  { value: "san_mateo", label: "San Mateo" },
  { value: "petaluma", label: "Petaluma" },
  { value: "french_valley", label: "French Valley" },
  { value: "burlingame", label: "Burlingame" },
  { value: "other", label: "Other / Not listed" },
];
const ACIES_ELECTRICAL_SCOPE_UTILITY_REGION_OPTIONS = [
  { value: "pge", label: "PG&E" },
  { value: "sce", label: "SCE" },
  { value: "sdge", label: "SDG&E" },
  { value: "iid", label: "IID" },
  { value: "smud", label: "SMUD" },
  { value: "ladwp", label: "LADWP" },
  { value: "pse", label: "PSE" },
  { value: "other", label: "Other / Not listed" },
];
const ACIES_ELECTRICAL_SCOPE_PROJECT_DELIVERY_OPTIONS = [
  { value: "tenant_improvement", label: "Tenant Improvement" },
  { value: "new_building", label: "New Building" },
  { value: "shell_core", label: "Shell / Core & Shell" },
  { value: "addition_remodel", label: "Addition / Remodel" },
  { value: "service_upgrade", label: "Service Upgrade" },
  { value: "other", label: "Other" },
];
const ACIES_ELECTRICAL_SCOPE_OCCUPANCY_OPTIONS = [
  { value: "office", label: "Office" },
  { value: "retail", label: "Retail" },
  { value: "restaurant", label: "Restaurant / Food Service" },
  { value: "warehouse", label: "Warehouse / Industrial" },
  { value: "grocery", label: "Grocery / Market" },
  { value: "mixed_use", label: "Mixed Use" },
  { value: "healthcare", label: "Healthcare" },
  { value: "other", label: "Other" },
];
const ACIES_ELECTRICAL_SCOPE_SERVICE_SIZE_OPTIONS = [
  { value: "lt400", label: "Under 400A" },
  { value: "400_999", label: "400A - 999A" },
  { value: "1000_1199", label: "1000A - 1199A" },
  { value: "1200_2999", label: "1200A - 2999A" },
  { value: "3000_plus", label: "3000A+" },
];
const WORKROOM_VERIFICATION_STATUS_OPTIONS = [
  { value: "not_started", label: "Not Started" },
  { value: "in_progress", label: "In Progress" },
  { value: "complete", label: "Complete" },
  { value: "not_available", label: "Not Available" },
];
const CHECKLIST_BINARY_SCOPE_OPTIONS = [
  { value: "yes", label: "Yes" },
  { value: "no", label: "No" },
];
const CHECKLIST_ISSUE_STAGE_OPTIONS = [
  { value: "permit", label: "Permit Set" },
  { value: "construction", label: "Construction Set" },
  { value: "revision", label: "Revision / Bulletin" },
  { value: "submittal", label: "Submittal / Coordination" },
  { value: "record", label: "Record / As-Built" },
  { value: "other", label: "Other" },
];
const WORKROOM_PHASE_OPTIONS = [
  {
    value: "pre_design",
    label: "Pre-Design",
    eyebrow: "Phase 1 of 4",
    description:
      "Define scope, capture project constraints, and prepare the CAD reference base before design starts.",
  },
  {
    value: "design",
    label: "Design",
    eyebrow: "Phase 2 of 4",
    description:
      "Work through the electrical design checklist, develop the permit set, and manage general notes.",
  },
  {
    value: "preflight",
    label: "Preflight & Issue",
    eyebrow: "Phase 3 of 4",
    description:
      "Run the issue checklist, tighten the drawing set, and prepare publishing actions for release.",
  },
  {
    value: "post_permit",
    label: "Post-Permit",
    eyebrow: "Phase 4 of 4",
    description:
      "Respond to plan check comments, service changes, and bulletin work after the initial permit set.",
  },
];
const WORKROOM_PHASE_VALUES = new Set(
  WORKROOM_PHASE_OPTIONS.map((phase) => phase.value)
);
const WORKROOM_RETURN_TYPE_OPTIONS = [
  { value: "", label: "Select return type..." },
  { value: "plan_check", label: "Plan Check" },
  { value: "service_change", label: "Service Change" },
  { value: "bulletin", label: "Bulletin / Revision" },
];
const WORKROOM_RETURN_TYPE_VALUES = new Set(
  WORKROOM_RETURN_TYPE_OPTIONS.map((option) => option.value).filter(Boolean)
);
const ACIES_ELECTRICAL_SCOPE_FIELD_CONFIG = [
  {
    key: "state",
    label: "State",
    helpText: "Used for state-specific code and energy workflow rows.",
    options: ACIES_ELECTRICAL_SCOPE_STATE_OPTIONS,
    section: "scope_jurisdiction",
  },
  {
    key: "cityArea",
    label: "City / Area",
    helpText: "Used for city- and area-specific notes inside the ACIES checklist.",
    options: ACIES_ELECTRICAL_SCOPE_CITY_AREA_OPTIONS,
    section: "scope_jurisdiction",
  },
  {
    key: "utilityRegion",
    label: "Utility Region",
    helpText: "Used for PG&E, SCE, SDG&E, IID, SMUD, LADWP, and PSE-specific checks.",
    options: ACIES_ELECTRICAL_SCOPE_UTILITY_REGION_OPTIONS,
    section: "scope_jurisdiction",
  },
  {
    key: "projectDeliveryType",
    label: "Scope / Work Type",
    helpText: "Used for tenant improvement, shell, service, and delivery-specific decisions.",
    options: ACIES_ELECTRICAL_SCOPE_PROJECT_DELIVERY_OPTIONS,
    section: "scope_jurisdiction",
  },
  {
    key: "occupancyFamily",
    label: "Building / Occupancy Type",
    helpText: "Used for office, kitchen, warehouse, retail, and grocery checks.",
    options: ACIES_ELECTRICAL_SCOPE_OCCUPANCY_OPTIONS,
    section: "building_profile",
  },
  {
    key: "buildingAreaSqFt",
    label: "Approx. Area (Sq Ft)",
    helpText: "Used for early sizing and scope conversations. Stored as a project-wide reference only.",
    inputType: "number",
    section: "building_profile",
    tokenizable: false,
  },
  {
    key: "storyCount",
    label: "Above Grade Stories",
    helpText: "Quick building scale reference for early design assumptions.",
    inputType: "number",
    section: "building_profile",
    tokenizable: false,
  },
  {
    key: "serviceSizeBucket",
    label: "Service Size",
    helpText: "Used for service-equipment checks that depend on ampacity thresholds.",
    options: ACIES_ELECTRICAL_SCOPE_SERVICE_SIZE_OPTIONS,
    section: "building_profile",
  },
  {
    key: "highRise",
    label: "High Rise Building",
    helpText: "Used for high-rise-specific notes.",
    options: CHECKLIST_BINARY_SCOPE_OPTIONS,
    section: "building_profile",
  },
  {
    key: "hasServiceEquipment",
    label: "Service Equipment / Switchboard",
    helpText: "Used for one-line and electrical-room service checks.",
    options: CHECKLIST_BINARY_SCOPE_OPTIONS,
    section: "existing_conditions",
  },
  {
    key: "hasTransformer",
    label: "Transformer",
    helpText: "Used for transformer section checks.",
    options: CHECKLIST_BINARY_SCOPE_OPTIONS,
    section: "existing_conditions",
  },
  {
    key: "hasExteriorSiteWork",
    label: "Exterior / Site Work",
    helpText: "Used for site lighting, utility, and pole/site infrastructure checks.",
    options: CHECKLIST_BINARY_SCOPE_OPTIONS,
    section: "existing_conditions",
  },
  {
    key: "hasRooftopMechanical",
    label: "Rooftop Mechanical / Roof Electrical",
    helpText: "Used for RTU, roof receptacle, and rooftop coordination checks.",
    options: CHECKLIST_BINARY_SCOPE_OPTIONS,
    section: "existing_conditions",
  },
  {
    key: "hasMeetingRooms",
    label: "Meeting / Conference Rooms",
    helpText: "Used for meeting-room receptacle checks.",
    options: CHECKLIST_BINARY_SCOPE_OPTIONS,
    section: "existing_conditions",
  },
  {
    key: "hasCommercialKitchen",
    label: "Commercial Kitchen / Food Prep",
    helpText: "Used for commercial kitchen GFCI and shunt-trip coordination checks.",
    options: CHECKLIST_BINARY_SCOPE_OPTIONS,
    section: "existing_conditions",
  },
  {
    key: "hasElevator",
    label: "Elevator",
    helpText: "Used for elevator-specific checklist sections.",
    options: CHECKLIST_BINARY_SCOPE_OPTIONS,
    section: "existing_conditions",
  },
  {
    key: "hasEmergencyPower",
    label: "Emergency Power / Generator",
    helpText: "Used for generator and emergency-power checks.",
    options: CHECKLIST_BINARY_SCOPE_OPTIONS,
    section: "existing_conditions",
  },
  {
    key: "hasFirePump",
    label: "Fire Pump",
    helpText: "Used for fire pump checks.",
    options: CHECKLIST_BINARY_SCOPE_OPTIONS,
    section: "existing_conditions",
  },
  {
    key: "hasSolar",
    label: "Solar / PV",
    helpText: "Used for solar-ready and PV-related checks.",
    options: CHECKLIST_BINARY_SCOPE_OPTIONS,
    section: "existing_conditions",
  },
  {
    key: "hasBatteryOrInverter",
    label: "Battery / Inverter System",
    helpText: "Used for inverter and storage-related checks.",
    options: CHECKLIST_BINARY_SCOPE_OPTIONS,
    section: "existing_conditions",
  },
  {
    key: "hasEvCharging",
    label: "EV Charging",
    helpText: "Used for EV charger and EV infrastructure checks.",
    options: CHECKLIST_BINARY_SCOPE_OPTIONS,
    section: "existing_conditions",
  },
  {
    key: "hasWarehouseLoading",
    label: "Warehouse / Loading Spaces",
    helpText: "Used for warehouse and loading-space EV provisions.",
    options: CHECKLIST_BINARY_SCOPE_OPTIONS,
    section: "existing_conditions",
  },
  {
    key: "surveyStatus",
    label: "Survey Report",
    helpText: "Track whether the site survey/report has been completed for this project.",
    options: WORKROOM_VERIFICATION_STATUS_OPTIONS,
    section: "field_verification",
    tokenizable: false,
  },
  {
    key: "sitePhotosStatus",
    label: "Site Photos",
    helpText: "Track whether field photos are available for design reference.",
    options: WORKROOM_VERIFICATION_STATUS_OPTIONS,
    section: "field_verification",
    tokenizable: false,
  },
  {
    key: "panelPhotosStatus",
    label: "Panel Photos",
    helpText: "Track whether panel and equipment photos have been gathered for design and AI tools.",
    options: WORKROOM_VERIFICATION_STATUS_OPTIONS,
    section: "field_verification",
    tokenizable: false,
  },
];
const ACIES_ELECTRICAL_SCOPE_FIELD_KEYS = ACIES_ELECTRICAL_SCOPE_FIELD_CONFIG.map(
  (field) => field.key
);
const CHECKLIST_ISSUE_STAGE_VALUES = new Set(
  CHECKLIST_ISSUE_STAGE_OPTIONS.map((option) => option.value)
);
const CHECKLIST_SCOPE_ALLOWED_VALUE_LOOKUP = ACIES_ELECTRICAL_SCOPE_FIELD_CONFIG.reduce(
  (lookup, field) => {
    if (Array.isArray(field.options) && field.options.length) {
      lookup[field.key] = new Set((field.options || []).map((option) => option.value));
    }
    return lookup;
  },
  {
    issueStage: CHECKLIST_ISSUE_STAGE_VALUES,
  }
);

function createChecklistTemplateSubheader(text) {
  return {
    type: CHECKLIST_ROW_TYPES.SUBHEADER,
    text: String(text ?? ""),
  };
}

function createChecklistTemplateItem(text, applicability = null) {
  const item = {
    type: CHECKLIST_ROW_TYPES.ITEM,
    text: String(text ?? ""),
  };
  const normalizedApplicability = normalizeChecklistApplicability(applicability);
  if (normalizedApplicability) {
    item.applicability = normalizedApplicability;
  }
  return item;
}

function createChecklistTemplateSection(title, items = []) {
  return [createChecklistTemplateSubheader(title), ...items];
}

const PRE_FLIGHT_ELECTRICAL_TEMPLATE_ITEMS = [
  createChecklistTemplateSubheader("General coordination"),
  "Check relevant codes that apply to the project based on local city, state, and national codes. (CEC 90.2, 90.4)",
  "Check if the tenant space has existing mechanical units on the roof that are not powered from tenant electrical panels. (CEC 430.102, 440.14)",
  "Check for unmentioned items that could need power (interior signage). (CEC 600.6)",
  "Check for occ sensor on ceiling instead of wall in storage room. (Title 24, Part 6 Section 130.1(c)1)",
  "Check for dedicated service receptacle within 25ft of electrical panels. (CEC 210.63(B)(2), 110.26(E))",
  "Check for junction box indicated as wall mount for hand dryers. (CEC 314.23, 314.29)",
  "Check for adequate space for electrical panels, relocate to BOH corridors as necessary, avoid storage rooms, IT server racks. (CEC 110.26(A), 110.26(E))",
  "Check for food waste disposer for all sinks with outlet under counter and switch above (confirm with plumbing as it is not a requirement). (CEC 422.16(B)(1), 422.31(B))",
  "Check for furniture systems and include note to verify point of connection for furniture systems. (CEC 605)",
  createChecklistTemplateSubheader("Power, receptacles & equipment"),
  "Check for controlled receptacles in office, lobby, kitchen, printer/copy room, conference room, meeting room. Modular furniture workstations need at least one controlled receptacle per workstation. (Title 24, Part 6 Section 130.5(d))",
  "Check for tamper proof receptacles in areas where children may be present: business offices, lobbies, waiting areas, theaters, auditoriums, gyms, bowling alleys, bus stations, airports, train stations. (CEC 406.12)",
  "Check for rooftop mechanical units shown on RCP, ensure they are dashed in appearance and noted to go on the roof along with rooftop receptacle. (CEC 210.63(A), 440.14)",
  "Check for return or supply air system over 2000CFM for duct smoke requirement. Any mechanical units 2000CFM or over should get duct smoke. (Title 24, Part 2 CBC Section 907.2.12.1.2)",
  "Check for hand dryer specification, if not add note to confirm exact breaker size with manufacturer. (CEC 110.3(B))",
  "Check for GFCI protection at required nonresidential receptacle locations (kitchens, outdoor, rooftops, within 6 ft of sinks, etc.). (CEC 210.8(B))",
  "Check meeting rooms for required receptacle outlets, including floor outlets when room size thresholds are met. (CEC 210.65)",
  createChecklistTemplateSubheader("Distribution & single-line"),
  "Check for kAIC rating shown on single line main switchboard. (CEC 110.9, 110.10)",
  "Check voltage drop requirements and calculations for feeders and branch circuits (<=5% combined). (Title 24, Part 6 Section 130.5(c))",
  "Check AIC ratings for panels and transformers on single line. (CEC 110.9, 110.10)",
  createChecklistTemplateSubheader("Lighting controls & Title 24"),
  "Check for >4000W in space needing demand response. Provide necessary software and device(s) to automatically reducing the lighting power by at least 15% upon receiving a demand response signal. (Title 24, Part 6 Section 110.12)",
  "Check for daylight harvesting in daylit zones. Provide daylight sensor & room controller for automatic dimming of light fixtures. (Title 24, Part 6 Section 130.1(d))",
  "Check for dimming in all rooms that are BOTH >100sqft AND >0.5W/sqft.",
  "Check for occupancy sensors which are required in offices, conference & meeting rooms, classrooms, restrooms, multipurpose rooms, warehouses, library aisles, corridors & stairwells.",
  "Check for time-switch controls which are allowed in lobbies, retail sales floors, commercial kitchens, auditoriums & theaters, and large multipurpose rooms.",
  "Check for 2-hour bypass switch in each room controlled by automatic time controlled on/off. Provide 1 bypass switch per 5000sqft of room size.",
  "Check for timeclock to indoor light fixtures that has manual override up to 2-hours, battery or internal memory capable of storing schedule for at least 7 days if power goes out.",
  "Check for astronomical timeclock, photosensor located on the roof, and contactors with signage, exterior / outdoor lights, pole lights, etc. (Section 130.2(c)2.B)",
  "Check for spaces that don't require multilevel controls such as: rooms under 100sqft, restrooms, rooms with only one luminaire.",
];

const ELECTRICAL_GENERAL_TEMPLATE_ITEMS = [
  createChecklistTemplateSubheader("General setup"),
  "Reference Manager",
  "\"cleanup\" command",
  "XREF BGs",
  "Sheet layout & Titleblock",
  "Scope of work",
  "Specification sheet (CA or other state-specific)",
  "Sheet Index",
  createChecklistTemplateSubheader("Distribution & major equipment"),
  "MSB / panelboards / meter",
  "SLD",
  "Panel schedules",
  "EV chargers",
  "Solar",
  createChecklistTemplateSubheader("Lighting plans & controls"),
  "Arch RCP notes",
  "Title 24",
  "Light fixture schedule (check site photos of existing lights and compare to archived bank standard)",
  "Light control schedule",
  "Out-of-scope",
  "Daylight zones",
  "Timeclock",
  "EM lights",
  "Symbol list",
  "Light power, controls, notes",
  createChecklistTemplateSubheader("Power plans & schedules"),
  "Arch power notes",
  "Out-of-scope",
  "Equipment schedule",
  "IG/GFI/WP/Controlled/General power outlets",
  "Symbol list",
  createChecklistTemplateSubheader("Site, roof & specialty systems"),
  "Solar equipment",
  "Solar meter",
  "HVAC equipment",
  "Roof outlets",
  "Solar zone",
  "Solar stub",
  "Heaters",
  "Pumps",
  "Lights",
  "Signs",
  "EV chargers",
  "Equipment",
  "Photometric",
];

const ACIES_MECHANICAL_TEMPLATE_ITEMS = [
  ...createChecklistTemplateSection("Mechanical General", [
    "Parapet height vs. RTU height + curb. Place all equipment at least 10 feet away from roof edge if the parapet is not at least 42\" tall.",
    "RTU voltage.",
    "R-32 or R-454B refrigerant.",
    "10' liner from unit for noise attenuation.",
    "Check EF noise level.",
    "Vertical vs. horizontal discharge.",
    "VAV - spec with transformer.",
    "Demand control ventilation / CO2 sensor requirement.",
    "Economizer requirement (no economizer for grocery store due to humidity).",
    "Economizer: barometric relief for 7.5-ton or less. Otherwise, provide power exhaust.",
    "Economizer: with horizontal discharge, barometric relief hood needs to be mounted on horizontal return ductwork.",
    "Smoke detector requirement. (CA & FL: SD on supply side).",
    "Smoke detector at supply fan and makeup air unit, 2000 CFM or greater.",
    "In-line pump for 4 pipe system - bid alternate.",
    "Latest state code requirement; year.",
    "Drawing schedule to match sheet numbers.",
    "Sheet name matches what is shown on drawing index.",
    "Note on control wiring; between split systems and to T-stats.",
    "Proper scale on drawings.",
    "Coordinate drawings presentation, sheet naming, sheet numbering, and scales with Electrical & Plumbing for consistency.",
    "Ventilation rate per occupancy.",
    "GFI for RTU: by EC or furnished with RTU.",
    "GFI receptacle to have dedicated neutral wire.",
    "RTU disconnect switch: by EC or furnished with RTU.",
    "RTU coating at coastal area.",
    "Any humidity control requirement?",
    "Hood control by whom?",
    "Duct insulation = R-8 (CA only).",
    "5' limitation on flex duct to diffuser.",
    "Provide elevator hoistway emergency venting if the driving machine is within the hoistway (CA only, check local building code).",
    "Smart T-stat specifications. (CA only) - Wi-Fi smart thermostat, T24 compliant, with demand response capability. Honeywell Mod. TH9320WF5003 Wi Fi 9000, or equal.",
    "For mini-split or VRF, specify 24V adapter interface.",
  ]),
  ...createChecklistTemplateSection("Mechanical Floor Plan", [
    "Number of units serving a single space.",
    "Ductwork available clearance.",
    "Ductwork clearance in mezzanine?",
    "Cooling in elevator machine room. Location of unit to be outside of EMR. Walls are usually rated.",
    "10' of liner from HVAC unit to reduce noise.",
    "Flue and combustion air for gas water heater.",
    "If multiple small units serving same space, provide smoke detectors if total CFM is over 2,000.",
    "Coordinate fire/smoke damper location with electrical.",
    "Fire wall protection @ shaft, electrical & mechanical room.",
    "Fire rating between floors.",
    "FSD & fire rated wall / ceiling.",
    "Duct detector on supply side (for all other states beside CA & FL, place on return side).",
    "Outside air provision for ground floor tenants.",
    "Toilet exhaust provision for ground floor tenants.",
    "T-stat location matching serving unit.",
    "Remote temp sensor in RTU + T-stat @ office.",
    "Ventilation for dimming panel or transformer.",
    "Electric heater in sprinkler room.",
    "Air curtain: cold climate zone, humidity?",
    "Split system secondary drain pan moisture sensor detail.",
    "FM requirement: smoke damper below EF.",
    "Provide anchor and guide spec for large wet pipe.",
    "Health care negative pressure room: 75 CFM or 10% more exhaust.",
    "Provide damper at louver locations.",
    "Provide bug screen for all exterior grilles/louvers. No screen for dryer exhaust.",
    "No exposed positive pressurized duct in plenum / open ceiling.",
  ]),
];

const ACIES_ELECTRICAL_TEMPLATE_ITEMS = [
  ...createChecklistTemplateSection("Electrical General", [
    createChecklistTemplateItem("Latest state code requirement; year.", {
      any: [
        "issueStage:permit",
        "issueStage:construction",
        "issueStage:revision",
      ],
    }),
    "Drawing schedule to match sheet numbers.",
    "Proper scale on drawings.",
    "Coordinate drawings presentation, sheet naming, sheet numbering, and scales with Mechanical & Plumbing for consistency.",
    "Lighting control diagram.",
    createChecklistTemplateItem("Lighting control specification.", {
      any: [
        "issueStage:permit",
        "issueStage:construction",
        "issueStage:revision",
      ],
    }),
    "Latest specification or applicable code.",
    createChecklistTemplateItem(
      "Non-infringement statement (for high rise building), and energy inspection form (for all projects) for City of San Francisco.",
      {
        all: ["cityArea:san_francisco"],
        any: [
          "issueStage:permit",
          "issueStage:construction",
          "issueStage:revision",
        ],
      }
    ),
    createChecklistTemplateItem(
      "Scope of work, CA codes list, and city stamp box (CSJ Bulletin #227) 6\" x 6\" cover sheet & 3\" x 3\" each sheet for City of San Jose.",
      {
        all: ["cityArea:san_jose"],
        any: [
          "issueStage:permit",
          "issueStage:construction",
          "issueStage:revision",
        ],
      }
    ),
    "Receptacle to be 18\" top of box, and switches to be 48\" top of box mounting height.",
    createChecklistTemplateItem("Utilize online T-24 forms.", {
      all: ["state:ca"],
      any: [
        "issueStage:permit",
        "issueStage:construction",
        "issueStage:revision",
      ],
    }),
    createChecklistTemplateItem("Utilize online energy forms for WA. https://waenergycodes.com/", {
      all: ["state:wa"],
      any: [
        "issueStage:permit",
        "issueStage:construction",
        "issueStage:revision",
      ],
    }),
    createChecklistTemplateItem("Utilize Comchecks for other states.", {
      none: ["state:ca", "state:wa"],
      any: [
        "issueStage:permit",
        "issueStage:construction",
        "issueStage:revision",
      ],
    }),
    "Check adopted energy code for lighting control requirements.",
    "Check municipal code for specific energy code requirements/amendments.",
    createChecklistTemplateItem("Check for Reach Code requirements.", {
      all: ["state:ca"],
    }),
  ]),
  ...createChecklistTemplateSection("Electrical Lighting", [
    "Sequence of operation.",
    "Light fixture voltage / dimmable / 0-10V.",
    "Light fixture tag match lighting plan.",
    "Track current limiter rating at 80% loading.",
    "Location of override switch for exterior lighting controls.",
    "Check night light (NL).",
    "ATM:AB244 lighting standard - 10FC at 5' x 5'; 2FC at 50' x 50'.",
    "Separate lighting control for daylit area.",
    "Power pack for low-voltage sensor.",
    "Location of photocell or photo-sensor.",
    "Lighting under stairway.",
    "Light switch (OS and manual on/off) location.",
    createChecklistTemplateItem("Any dimming control requirement for other states.", {
      none: ["state:ca"],
    }),
    "Lighting above all loading door.",
    "Separate switching for control access area (Pharmacy).",
    "Exterior light specification, UL listed for wet location. BUG rating.",
    "Any dimming ballast requirement.",
    "New stairway lighting 10fc min not ave., NFPA 101.",
  ]),
  ...createChecklistTemplateSection("HVAC coord", [
    createChecklistTemplateItem(
      "Provide motion sensor and light in attic space / near roof ladder opening, CMC 304.3.2.",
      {
        all: ["hasRooftopMechanical:yes"],
      }
    ),
    "Toilet EF control. Interlock?",
  ]),
  ...createChecklistTemplateSection("T24", [
    createChecklistTemplateItem(
      "Dimming control for >=100SF with general lighting 0.5w/sf per T24.",
      {
        all: ["state:ca"],
      }
    ),
    createChecklistTemplateItem(
      "Demand response design for lighting power 4000W subject to dimming requirements per T24 110.12(c) and controlled receptacles per T24 110.12(e).",
      {
        all: ["state:ca"],
      }
    ),
  ]),
  ...createChecklistTemplateSection("Elevator", [
    createChecklistTemplateItem("Elevator pit - use 4' / 8' LED. 10fc minimum.", {
      all: ["hasElevator:yes"],
    }),
    createChecklistTemplateItem(
      "Elevator lobby (landing) to maintain 10FC per ASME A17.1:2004 2.1110.2.",
      {
        all: ["hasElevator:yes"],
      }
    ),
  ]),
  ...createChecklistTemplateSection("EM Lighting", [
    "EM lighting (interior & exterior) per CBC 1008.",
    "EM lighting relays, UL924 to bypass lighting controls.",
    "EM light circuiting per 700.12.",
    createChecklistTemplateItem(
      "EM lighting - check if high rise building has generator power for emergency light.",
      {
        all: ["highRise:yes"],
      }
    ),
    "Notes on EM circuit ahead of switches.",
    "Sufficient EM lighting. 1fc ave. and min. at any point of 0.1fc per CBC 1008.3.5.",
    "Exterior EM light above exit doors.",
    createChecklistTemplateItem(
      "County of Santa Clara: provide night light for egress lighting in every room under normal conditions.",
      {
        all: ["cityArea:santa_clara_county"],
      }
    ),
  ]),
  ...createChecklistTemplateSection("Electrical Site", [
    createChecklistTemplateItem("Exterior lighting average 1fc min 0.1fc - at egress path.", {
      all: ["hasExteriorSiteWork:yes"],
    }),
    createChecklistTemplateItem("Junction box every 200'.", {
      all: ["hasExteriorSiteWork:yes"],
    }),
    "Demarcation location and conduits. (2)4\"C.",
    createChecklistTemplateItem("PIV location.", {
      all: ["hasExteriorSiteWork:yes"],
    }),
    createChecklistTemplateItem("EV charger location - CALGREEN.", {
      all: ["state:ca", "hasEvCharging:yes"],
    }),
    "Signage location. Monument sign? Illuminated address sign.",
    createChecklistTemplateItem("Any sump pump at loading dock?", {
      all: ["hasWarehouseLoading:yes"],
    }),
    createChecklistTemplateItem("IID details is different from PG&E & SCE.", {
      all: ["utilityRegion:iid"],
    }),
    createChecklistTemplateItem("SDG&E: 6' clearance in front of pull section.", {
      all: ["utilityRegion:sdge"],
    }),
    createChecklistTemplateItem("Landscape/irrigation controller.", {
      all: ["hasExteriorSiteWork:yes"],
    }),
    createChecklistTemplateItem("Any irrigation pump?", {
      all: ["hasExteriorSiteWork:yes"],
    }),
    createChecklistTemplateItem("Any sewer or grinder pump? simplex vs. duplex? EYS fitting?", {
      all: ["hasExteriorSiteWork:yes"],
    }),
    createChecklistTemplateItem("Pole height limitation by city.", {
      all: ["hasExteriorSiteWork:yes"],
    }),
    createChecklistTemplateItem("Pole light 24' or lower shall have dimming (CA only).", {
      all: ["state:ca", "hasExteriorSiteWork:yes"],
    }),
    createChecklistTemplateItem("Any receptacle on light pole.", {
      all: ["hasExteriorSiteWork:yes"],
    }),
    "Coordination POC with JT consultant.",
    createChecklistTemplateItem("CCT (color correction temperature) limitation by city.", {
      all: ["hasExteriorSiteWork:yes"],
    }),
    createChecklistTemplateItem("Full cutoff for all exterior lighting limitations by city.", {
      all: ["hasExteriorSiteWork:yes"],
    }),
    createChecklistTemplateItem("Bi-level lighting control.", {
      all: ["hasExteriorSiteWork:yes"],
    }),
    "Distance from utility transformer to window, door & swbd.",
    "Coordination POC with JT consultant.",
  ]),
  ...createChecklistTemplateSection("Electrical Power", [
    "Review electrical power requirements and coordination notes.",
  ]),
  ...createChecklistTemplateSection("EL room", [
    "Clearance in front of panels and switchboard.",
    createChecklistTemplateItem("2\"C PG&E conduit from switchgear to exterior wall.", {
      all: ["hasServiceEquipment:yes", "utilityRegion:pge"],
    }),
    createChecklistTemplateItem("PG&E: 48\" clearance in front of switchboard.", {
      all: ["hasServiceEquipment:yes", "utilityRegion:pge"],
    }),
    createChecklistTemplateItem(
      "New Mexico: disconnect means on outside of bldg. shall be within 48\" from service conductors emerge from earth.",
      {
        all: ["hasServiceEquipment:yes", "state:nm"],
      }
    ),
    "Space for (side or rear) pull section at switchboard.",
    "2 exit doors for equipment 1200A or above and more than 6' wide; door to swing out and with panic hardware for 800A & above.",
    "Provide receptacle within 15' of MSB.",
    "6\" wall for recessed panel.",
    "No ductwork in EL room.",
    "SPD for residential/emergency, see EATON guidelines.",
    "Fire rating of room? - if emergency non-sprinklered - 2hr.",
    createChecklistTemplateItem(
      "Solar and battery storage requirement per T24. Use performance approach to avoid battery storage system.",
      {
        all: ["state:ca"],
      }
    ),
    "Solar ready - conduits to roof, solar PV space.",
    "FACP?",
    "MPOE - 4' board.",
    createChecklistTemplateItem(
      "SCE does not allow fire/security alarm, batteries and irrigation controllers inside meter room.",
      {
        all: ["utilityRegion:sce"],
      }
    ),
    createChecklistTemplateItem(
      "SDG&E does not allow telephone, data, communication, fire/security alarm, irrigation control and non-SDG&E related equipment in meter room.",
      {
        all: ["utilityRegion:sdge"],
      }
    ),
  ]),
  ...createChecklistTemplateSection("HVAC/Plumbing coord", [
    createChecklistTemplateItem("LG is most efficient for RTU/VRF. (Load).", {
      all: ["hasRooftopMechanical:yes"],
    }),
    "Any HVAC/plumbing pipe above panels?",
    createChecklistTemplateItem("3/4\"C., between FC & CU.", {
      all: ["hasRooftopMechanical:yes"],
    }),
    createChecklistTemplateItem("GFI for RTU: by EC or furnished with RTU.", {
      all: ["hasRooftopMechanical:yes"],
    }),
    createChecklistTemplateItem("Disconnect switch for RTU: by EC or furnished with RTU.", {
      all: ["hasRooftopMechanical:yes"],
    }),
    createChecklistTemplateItem("Coordinate damper(FSD), smoke detector/DD location.", {
      all: ["hasRooftopMechanical:yes"],
    }),
    createChecklistTemplateItem("Shunt-trip breakers for equipment under hood.", {
      all: ["hasCommercialKitchen:yes"],
    }),
    createChecklistTemplateItem("Shunt-trip breakers for Big Ass fan in WH.", {
      any: ["occupancyFamily:warehouse", "hasWarehouseLoading:yes"],
    }),
    createChecklistTemplateItem("Equipment with >=1/2HP motor, use disconnect switch.", {
      all: ["hasRooftopMechanical:yes"],
    }),
    createChecklistTemplateItem("Motor starter for large EF.", {
      all: ["hasRooftopMechanical:yes"],
    }),
    "Coordinate EWH voltage and balanced load per phase requirement.",
    createChecklistTemplateItem("Interlocking kitchen EF with MAU.", {
      all: ["hasCommercialKitchen:yes"],
    }),
  ]),
  ...createChecklistTemplateSection("Generator/Inverter", [
    createChecklistTemplateItem("Circuiting to genset chargers and jacket water heater.", {
      all: ["hasEmergencyPower:yes"],
    }),
    createChecklistTemplateItem(
      "EM generator: permanent switching means to connect a temporary source of power for maintenance or repair of the alternate source of power in accordance with 700.3(F).",
      {
        all: ["hasEmergencyPower:yes"],
      }
    ),
  ]),
  ...createChecklistTemplateSection("Fire Pump", [
    createChecklistTemplateItem("Must come with jockey pump. All WP inside pump room.", {
      all: ["hasFirePump:yes"],
    }),
  ]),
  ...createChecklistTemplateSection("General", [
    "Show window receptacles per 210.62.",
    "Signage location vs. architectural layout.",
    "Any structural expansion joint?",
    "Power & Tel conduit stub to ceiling space for multi story bldg.",
    createChecklistTemplateItem("Receptacle requirement for meeting rooms (conference rooms) per 210.65.", {
      all: ["hasMeetingRooms:yes"],
    }),
    "Tamper resistant (406.12), hospital grade?",
    "GFCI protection for receptacles for counter and near water source (6').",
    createChecklistTemplateItem(
      "GFCI protection for restaurant kitchens or areas with a sink and permanent provisions for either food preparation or cooking per 210.8(B)(2).",
      {
        all: ["hasCommercialKitchen:yes"],
      }
    ),
    "Weatherproof in-use type and weather-resistant devices for exterior receptacles.",
    "GFCI protection for indoor service equipment and indoor equipment requiring dedicated equipment space per 210.8(E)/210.63.",
    "Garage level - provide GFI breakers for heat trace snow melt equipment.",
    "Any isolated ground requirement.",
    createChecklistTemplateItem(
      "Offices T24 CA: control receptacles for TI with new electrical service. No control receptacle is required for regular TI.",
      {
        all: ["state:ca", "occupancyFamily:office"],
      }
    ),
    "Any area of refuge?",
    "Separate structure (trailer, trash enclosure, pylon sign) with more than 1 branch circuit requires grounding electrodes, NEC 250.32 - should have ground rod + bonding to steel/cold water main.",
    "Avoid 3 hour rated wall, CBC 706.2.",
    createChecklistTemplateItem(
      "Shell development: all conduits shall maintain 5' clearance in front of HVAC supply and return openings.",
      {
        all: ["projectDeliveryType:shell_core"],
      }
    ),
    createChecklistTemplateItem(
      "Medium- and heavy-duty EV charging provision for warehouses, grocery stores and retail stores with planned off-street loading spaces per CGBC 5.106.5.4.",
      {
        all: ["state:ca", "hasWarehouseLoading:yes"],
        any: [
          "occupancyFamily:warehouse",
          "occupancyFamily:grocery",
          "occupancyFamily:retail",
        ],
      }
    ),
  ]),
  ...createChecklistTemplateSection("City/Area Specifics", [
    createChecklistTemplateItem("Milpitas - receptacle requirement.", {
      all: ["cityArea:milpitas"],
    }),
    createChecklistTemplateItem(
      "City of San Mateo: justify 1kW solar for residential; 5kW solar for commercial.",
      {
        all: ["cityArea:san_mateo"],
      }
    ),
    createChecklistTemplateItem(
      "City of Petaluma - mitigation measure - exterior receptacles & 220/1P receptacle at interior loading dock.",
      {
        all: ["cityArea:petaluma"],
      }
    ),
  ]),
  ...createChecklistTemplateSection("Electrical 1-Line Diagram", [
    createChecklistTemplateItem("EUSERC needs a switchboard section, no wireway.", {
      all: ["hasServiceEquipment:yes"],
    }),
    createChecklistTemplateItem("MCB for 6 or more services.", {
      all: ["hasServiceEquipment:yes"],
    }),
    "Upsize bus for PV - max PV size is 20% over bus size.",
    createChecklistTemplateItem(
      "Two to six service disconnects to have separate enclosures or separate vertical section per 230.71(B).",
      {
        all: ["hasServiceEquipment:yes"],
      }
    ),
    createChecklistTemplateItem("GFI @ MCB for >=1000A, 480V. (230.95).", {
      all: ["hasServiceEquipment:yes"],
      any: [
        "serviceSizeBucket:1000_1199",
        "serviceSizeBucket:1200_2999",
        "serviceSizeBucket:3000_plus",
      ],
    }),
    createChecklistTemplateItem(
      ">=1200 breaker: breaker with arc flash reduction maintenance system. (240.87).",
      {
        all: ["hasServiceEquipment:yes"],
        any: [
          "serviceSizeBucket:1200_2999",
          "serviceSizeBucket:3000_plus",
        ],
      }
    ),
    "CT for meter > 200A service. (full section).",
    createChecklistTemplateItem("Side or rear pull section at swbd. Top/bottom conduit opening.", {
      all: ["hasServiceEquipment:yes"],
    }),
    createChecklistTemplateItem("Any bus duct requirement by utility company.", {
      all: ["hasServiceEquipment:yes"],
    }),
    createChecklistTemplateItem("3000A - PG&E, bus trench 3000A LADWP.", {
      all: ["hasServiceEquipment:yes", "serviceSizeBucket:3000_plus"],
      any: ["utilityRegion:pge", "utilityRegion:ladwp"],
    }),
    createChecklistTemplateItem("LADWP - methane? Check Zimas.", {
      all: ["utilityRegion:ladwp"],
    }),
    "Disaggregation details.",
    "Voltage drop table.",
    createChecklistTemplateItem(
      "AIC rating and series rating. Series rating will require manufacturer catalogue.",
      {
        all: ["hasServiceEquipment:yes"],
      }
    ),
    createChecklistTemplateItem("SPD.", {
      all: ["hasServiceEquipment:yes"],
    }),
    createChecklistTemplateItem("Feeder taps overcurrent protection per 240.21(B) (identify distance).", {
      all: ["hasServiceEquipment:yes"],
    }),
    "240V delta service needs cable color identification and nameplate. (408.3(F), 110.21(B), 110.15).",
  ]),
  ...createChecklistTemplateSection("Panelboards", [
    "Panel board AIC rating, NEMA rating.",
    "Provide breaker handle-ties as required.",
    createChecklistTemplateItem(
      "Provide GFI breaker and lock-on breaker as required. EVCS without disconnect needs 110.25.",
      {
        all: ["hasEvCharging:yes"],
      }
    ),
    "Any isolated ground requirement.",
  ]),
  ...createChecklistTemplateSection("Elevator", [
    createChecklistTemplateItem("Shunt trip for elevator connections.", {
      all: ["hasElevator:yes"],
    }),
  ]),
  ...createChecklistTemplateSection("Grounding", [
    "Must have ground rod. Ground to cold water main/steel.",
    "Check grounding in mixed used bldg., oversize ground wire on 277/480V side.",
  ]),
  ...createChecklistTemplateSection("Transformer", [
    createChecklistTemplateItem(
      "Overcurrent protection per 450.3(B) - <1000V, primary only OK. Primary disconnect can be remote with identification 450.14.",
      {
        all: ["hasTransformer:yes"],
      }
    ),
    createChecklistTemplateItem("Transformer secondary overcurrent protection req per 240.21(C).", {
      all: ["hasTransformer:yes"],
    }),
    createChecklistTemplateItem("Feeder taps overcurrent protection per 240.21(B) (identify distance).", {
      all: ["hasTransformer:yes"],
    }),
    createChecklistTemplateItem("Grounding 250.30, 250.104.", {
      all: ["hasTransformer:yes"],
    }),
    createChecklistTemplateItem("Fire rating 1hr >112.5kVA if class <150.", {
      all: ["hasTransformer:yes"],
    }),
  ]),
  ...createChecklistTemplateSection("Generator/Battery/Inverter", [
    createChecklistTemplateItem(
      "Life safety generator to have portable generator docking station, NEC 700.3(F) - ATS has requirements.",
      {
        any: ["hasEmergencyPower:yes", "hasBatteryOrInverter:yes"],
      }
    ),
    createChecklistTemplateItem("Identify load type: life safety/legal/optional standby.", {
      any: ["hasEmergencyPower:yes", "hasBatteryOrInverter:yes"],
    }),
    createChecklistTemplateItem("ATS has requirements.", {
      any: ["hasEmergencyPower:yes", "hasBatteryOrInverter:yes"],
    }),
    createChecklistTemplateItem("Inverter sizing - maximum load at 80%. Show kWH.", {
      any: ["hasEmergencyPower:yes", "hasBatteryOrInverter:yes"],
    }),
    createChecklistTemplateItem("SPD.", {
      any: ["hasEmergencyPower:yes", "hasBatteryOrInverter:yes"],
    }),
    createChecklistTemplateItem("Fire code: clearances, fire alarm, 2hr rating. Check CFC 1207.7.", {
      any: ["hasEmergencyPower:yes", "hasBatteryOrInverter:yes"],
    }),
    createChecklistTemplateItem("NFPA 110 for room requirements.", {
      any: ["hasEmergencyPower:yes", "hasBatteryOrInverter:yes"],
    }),
  ]),
  ...createChecklistTemplateSection("Fire Pump", [
    createChecklistTemplateItem("Fire pump - meter - no main per NEC.", {
      all: ["hasFirePump:yes"],
    }),
    createChecklistTemplateItem("Fire pump - check reliable utility or req. generator.", {
      all: ["hasFirePump:yes"],
    }),
    createChecklistTemplateItem(
      "Fire pump = life safety; need separation barrier for breaker at generator/can't be combined.",
      {
        all: ["hasFirePump:yes"],
      }
    ),
  ]),
  ...createChecklistTemplateSection("City/Area Specifics", [
    createChecklistTemplateItem("Outdoor meters - Sacramento SMUD, OR, WA.", {
      any: ["utilityRegion:smud", "state:or", "state:wa"],
    }),
    createChecklistTemplateItem("SMUD - must have MCB, outdoor meters.", {
      all: ["utilityRegion:smud"],
    }),
    createChecklistTemplateItem("French Valley, CA requires house panel for single tenant pad bldg.", {
      all: ["cityArea:french_valley"],
    }),
    createChecklistTemplateItem("No shunt trip for elevator connections in Burlingame.", {
      all: ["cityArea:burlingame", "hasElevator:yes"],
    }),
    createChecklistTemplateItem("SFEC requires purple for high leg for 240V delta service. (SFEC 210(A)(3)).", {
      all: ["cityArea:san_francisco"],
    }),
    createChecklistTemplateItem("WA State - outdoor meters.", {
      all: ["state:wa"],
    }),
    createChecklistTemplateItem("WA State - separate meter for HVAC if over 50,000 SQFT.", {
      all: ["state:wa"],
    }),
    createChecklistTemplateItem("WA State - extra clearance for swbd - PSE.", {
      all: ["state:wa", "utilityRegion:pse"],
    }),
    createChecklistTemplateItem(
      "WA State - shunt trip at disconnect switch located within elevator machine room, load center in elevator machine room for ltg & rec; FC disconnect inside elevator machine room.",
      {
        all: ["state:wa", "hasElevator:yes"],
      }
    ),
  ]),
];

const ACIES_PLUMBING_TEMPLATE_ITEMS = [
  ...createChecklistTemplateSection("Plumbing General", [
    "General notes to be county or city specific.",
    "Check state code if seismic restraints are required for plumbing.",
    "Check if the project site falls under the IECC code.",
    "Check the current code being followed by the city.",
    "No gas service for Berkeley (low-rise & single family) starting 01/01/2020.",
    "Natural gas piping is not allowed for new ground up projects in California.",
    "Coordinate EWH voltage requirement.",
    "County of Santa Clara - submeter and remote reader compatibility.",
    "ADA compliance note.",
    "Drawing schedule to match sheet numbers.",
    "Pipe material specification.",
    "Note on access panel for clean out.",
    "Gas regulator.",
    "Proper scale on drawings.",
    "LADBS, LA health dept. and Irvine: minimum 30 gallons WH for all non-residential projects.",
    "Imperial County Health Dept: 3KW water heater: 25 GPH with 50 degree rise.",
    "Lavatory spec.",
    "Flush valve - manual / battery / power?",
    "Hose bib rating specification - anti-freeze?",
    "Fixture / faucets to comply with CalGreen and CPC for maximum flow rate or GPF requirements.",
    "FOG requirement by EBMUD - Alameda, Albany, Berkeley, Emeryville, Oakland & etc.",
    "Floor drain required when a bathroom has more than 1 water closet and / or with urinal.",
    "Drawing text size & width must be uniform and legible to read.",
    "No underground gas piping within the building footprint.",
    "Check the maximum fixture unit loading for horizontal vent piping.",
    "Avoid structural footings.",
    "Coordinate drawings presentation, sheet naming, sheet numbering, and scales with Electrical & Mechanical for consistency.",
  ]),
  ...createChecklistTemplateSection("Plumbing", [
    "Re-check the HVAC equipment schedule and layout prior to final submittal.",
    "Check gas meter location and capacity.",
    "Check all POC's for sewer, vent, water and gas.",
    "For projects involving flushvalve fixtures, check water pressure.",
    "Any sewer or grinder pump? simplex vs. duplex?",
    "No cast iron on pumped drain line. Provide either PVC or copper DWV (if PVC is not permitted).",
    "Check final sewer invert elevations.",
    "Rainwater discharge - underground or daylight per coordination.",
    "Cleanout every 100' or 50' if it is 1% slope.",
    "Provide cleanout at the base of vertical leader for RWL connected to the underground storm drain system.",
    "Avoid lengthy overhead condensate piping.",
    "Riverside County Health Dept: must have FS for case cooler.",
    "Floor sink must be accessible. Half-in half-out or directly under a sink with long legs.",
    "Provide drain at fire riser room, and check if standpipe is needed.",
    "Avoid any overhead plumbing lines directly above electrical room or server room.",
    "Access & clearance around water heater.",
    "Check ceiling height before specifying a wall mount platform for water heater above mop sink.",
    "Corrosive soil? ABS vs cast iron. Check pipe material table.",
    "Provide GI/GT for 3-comp sink.",
    "GI vent piping connections.",
    "GI manhole covers must be equal to the number of baffles.",
    "Clearance around WH.",
    "Water hammer arresters.",
    "Gas pipe in wall.",
    "Add roof mount condensate drain support detail.",
    "Canopy drains.",
    "No cleanout allowed in machine room.",
    "Flush type vs. tank type symbol.",
    "Small BFP for appliances.",
    "Sufficient number of sewer riser for upper floor connection, avoid slope.",
    "Vent provision for ground floor tenants for multi-story bldg. projects.",
    "Drainage for trash enclosure. Hose bibb, trap primer and GT.",
    "Pipe insulation.",
    "Storm drains: use same size for vertical and horizontal run.",
    "Hose bib location, on roof?",
    "HW & CW requirement at trash enclosure location.",
    "Water pressure and booster pump requirement.",
    "Main PRV location (outdoor by civil, indoor by PL).",
    "Sump pump at loading dock.",
    "# of elevator pits.",
    "Any water sub-meter requirement for TI space?",
    "Exterior pipe insulation.",
    "City of Santa Clara - single level type faucet.",
    "CCSF - backflow preventers on both hot & cold domestic water lines for the food disposal in the dishwashing room.",
    "Central Contra Costa Sanitary District - no condensate drain to SS.",
    "Backwater valve and sensor required in Burlingame.",
    "Shared grease interceptor allowed by city?",
    "County of Sacramento - no shared grease interceptor.",
    "City of Palo Alto - backwater valve for ground floor tenant.",
    "Clean out above urinal - CPC 707.4.",
    "Daly City - no underground PVC pipe.",
    "Provide anchor and guide spec for large wet pipe.",
    "All gas pipe to have seismic shut-off valves.",
    "Shell development: all piping shall maintain 5' clearance in front of HVAC supply and return openings.",
    "San Diego - no condensate drain into mop sink. Not also in the landscape; drywell ok.",
    "3-comp sink direct connection or indirect into FS.",
    "Avoid ADA parking stalls for rawwater discharge at face of curb.",
    "Roof drains directly above electrical room?",
    "Avoid hose bibb near narrow framed wall.",
    "Coordinate 4\" or larger vertical pipes with architect.",
    "Shell projects: RTU's to be connected to gas or just stub gas on roof?",
    "Verify water pressure especially for new development of buildings.",
    "Coordinate any electrical or panels with electrical team.",
    "Coordinate flue exhaust with mechanical team.",
    "Check with civil on storm drain to catch basin? Basin may be located on opposite side of roof drain (San Leandro Village).",
    "Check VTR.",
    "Avoid 3 hour rated wall, CBC 706.2.",
    "Garage / basement level / podium level - concrete apron around cleanout, AD.",
    "Sand / oil separator & grease interceptor - provide vent pipe.",
    "2 valves at IHW per T-24, CA only.",
    "Always check foundation footings & shear walls.",
    "Coordinate fan coil units with mechanical quantity and pump requirement.",
    "Avoid running overhead piping across the roof hatch/ access.",
    "Plumbing details to be job specific.",
    "Avoid floor mount toilets or floor sinks near the perimeter walls due to possible footings.",
    "Wall hung flush valve toilet - check if there is sufficient wall clearance / thickness. Typically requires:",
  ]),
];

const CHECKLIST_TEMPLATES = [
  {
    key: CHECKLIST_TEMPLATE_KEYS.PRE_FLIGHT_ELECTRICAL,
    name: PRE_FLIGHT_ELECTRICAL_CHECKLIST_NAME,
    discipline: "Electrical",
    items: PRE_FLIGHT_ELECTRICAL_TEMPLATE_ITEMS,
  },
  {
    key: CHECKLIST_TEMPLATE_KEYS.ELECTRICAL_GENERAL,
    name: ELECTRICAL_GENERAL_CHECKLIST_NAME,
    discipline: "Electrical",
    items: ELECTRICAL_GENERAL_TEMPLATE_ITEMS,
  },
  {
    key: CHECKLIST_TEMPLATE_KEYS.ACIES_MECHANICAL,
    name: ACIES_MECHANICAL_CHECKLIST_NAME,
    discipline: "Mechanical",
    alwaysVisible: true,
    items: ACIES_MECHANICAL_TEMPLATE_ITEMS,
  },
  {
    key: CHECKLIST_TEMPLATE_KEYS.ACIES_ELECTRICAL,
    name: ACIES_ELECTRICAL_CHECKLIST_NAME,
    discipline: "Electrical",
    alwaysVisible: true,
    items: ACIES_ELECTRICAL_TEMPLATE_ITEMS,
  },
  {
    key: CHECKLIST_TEMPLATE_KEYS.ACIES_PLUMBING,
    name: ACIES_PLUMBING_CHECKLIST_NAME,
    discipline: "Plumbing",
    alwaysVisible: true,
    items: ACIES_PLUMBING_TEMPLATE_ITEMS,
  },
];

function getBundledChecklistTemplateByKey(templateKey) {
  if (!templateKey) return null;
  return (
    CHECKLIST_TEMPLATES.find(
      (template) => template.key === String(templateKey).trim()
    ) || null
  );
}

function normalizeChecklistTemplateOverride(rawOverride, templateKey) {
  const baseTemplate = getBundledChecklistTemplateByKey(templateKey);
  if (!baseTemplate) {
    return { changed: false, override: null };
  }

  if (!rawOverride || typeof rawOverride !== "object" || Array.isArray(rawOverride)) {
    return { changed: rawOverride != null, override: null };
  }

  if (!Array.isArray(rawOverride.items)) {
    return { changed: true, override: null };
  }

  let changed = false;
  const rawName = typeof rawOverride.name === "string" ? rawOverride.name : "";
  const name = rawName.trim() || baseTemplate.name;
  if (name !== rawName) changed = true;

  const items = rawOverride.items.map((item, index) => {
    const normalizedItem = normalizeChecklistItem(item, index);
    changed = changed || normalizedItem.changed;
    const snapshot = {
      text: normalizedItem.item.text,
      type: normalizedItem.item.type,
    };
    if (normalizedItem.item.applicability) {
      snapshot.applicability = normalizedItem.item.applicability;
    }
    return snapshot;
  });

  return {
    changed,
    override: {
      name,
      items,
      updatedAt:
        typeof rawOverride.updatedAt === "string" ? rawOverride.updatedAt : null,
    },
  };
}

function getChecklistTemplateByKey(templateKey) {
  const baseTemplate = getBundledChecklistTemplateByKey(templateKey);
  if (!baseTemplate) return null;

  const override =
    checklistsDb?.templateOverrides &&
    typeof checklistsDb.templateOverrides === "object" &&
    !Array.isArray(checklistsDb.templateOverrides)
      ? checklistsDb.templateOverrides[baseTemplate.key]
      : null;

  if (!override) {
    return baseTemplate;
  }

  return {
    ...baseTemplate,
    name:
      typeof override.name === "string" && override.name.trim()
        ? override.name.trim()
        : baseTemplate.name,
    items: Array.isArray(override.items) ? override.items : baseTemplate.items,
  };
}

function getChecklistTemplatesForCurrentDisciplines() {
  const selectedDisciplines = new Set(
    normalizeDisciplineList(userSettings?.discipline).map((discipline) =>
      String(discipline || "").trim().toLowerCase()
    )
  );
  return CHECKLIST_TEMPLATES.filter((template) => {
    if (template?.alwaysVisible === true) return true;
    return selectedDisciplines.has(String(template.discipline || "").toLowerCase());
  })
    .map((template) => getChecklistTemplateByKey(template.key))
    .filter(Boolean);
}

function normalizeChecklistScopeOptionValue(fieldKey, value) {
  const allowedValues = CHECKLIST_SCOPE_ALLOWED_VALUE_LOOKUP[fieldKey];
  const raw = String(value || "").trim().toLowerCase();
  if (!raw || !allowedValues || !allowedValues.has(raw)) return "";
  return raw;
}

function getChecklistScopeFieldConfig(fieldKey) {
  return (
    ACIES_ELECTRICAL_SCOPE_FIELD_CONFIG.find((field) => field.key === fieldKey) || null
  );
}

function normalizeChecklistScopeNumberValue(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return "";
  const compact = raw.replace(/,/g, "");
  if (!/^\d+(\.\d+)?$/.test(compact)) return "";
  const normalized = compact.replace(/^0+(?=\d)/, "");
  return normalized || "0";
}

function normalizeChecklistScopeFieldValue(fieldKey, value) {
  const field = getChecklistScopeFieldConfig(fieldKey);
  if (!field) return "";
  if (field.inputType === "number") {
    return normalizeChecklistScopeNumberValue(value);
  }
  return normalizeChecklistScopeOptionValue(fieldKey, value);
}

function createEmptyAciesElectricalChecklistAnswers() {
  return ACIES_ELECTRICAL_SCOPE_FIELD_CONFIG.reduce((answers, field) => {
    answers[field.key] = "";
    return answers;
  }, {});
}

function normalizeAciesElectricalChecklistAnswers(answers) {
  const source = answers && typeof answers === "object" && !Array.isArray(answers)
    ? answers
    : {};
  const normalized = createEmptyAciesElectricalChecklistAnswers();
  ACIES_ELECTRICAL_SCOPE_FIELD_KEYS.forEach((fieldKey) => {
    normalized[fieldKey] = normalizeChecklistScopeFieldValue(fieldKey, source[fieldKey]);
  });
  return normalized;
}

function createDefaultAciesElectricalChecklistProfile() {
  return {
    schemaVersion: CHECKLIST_PROFILE_SCHEMA_VERSION,
    answers: createEmptyAciesElectricalChecklistAnswers(),
    updatedAtUtc: "",
  };
}

function normalizeAciesElectricalChecklistProfile(profile) {
  const source = profile && typeof profile === "object" && !Array.isArray(profile)
    ? profile
    : {};
  return {
    schemaVersion: CHECKLIST_PROFILE_SCHEMA_VERSION,
    answers: normalizeAciesElectricalChecklistAnswers(source.answers),
    updatedAtUtc:
      typeof source.updatedAtUtc === "string" ? source.updatedAtUtc.trim() : "",
  };
}

function normalizeProjectChecklistProfiles(profiles) {
  const source = profiles && typeof profiles === "object" && !Array.isArray(profiles)
    ? profiles
    : {};
  return {
    ...source,
    [ACIES_ELECTRICAL_SCOPE_PROFILE_KEY]: normalizeAciesElectricalChecklistProfile(
      source[ACIES_ELECTRICAL_SCOPE_PROFILE_KEY] || createDefaultAciesElectricalChecklistProfile()
    ),
  };
}

function normalizeChecklistIssueStage(value) {
  const raw = String(value || "").trim().toLowerCase();
  return CHECKLIST_ISSUE_STAGE_VALUES.has(raw) ? raw : "";
}

function normalizeWorkroomPhase(value) {
  const raw = String(value || "").trim().toLowerCase();
  return WORKROOM_PHASE_VALUES.has(raw) ? raw : "pre_design";
}

function normalizeWorkroomReturnType(value) {
  const raw = String(value || "").trim().toLowerCase();
  return WORKROOM_RETURN_TYPE_VALUES.has(raw) ? raw : "";
}

function createDefaultChecklistContext() {
  return {
    issueStage: "",
  };
}

function normalizeChecklistContext(context) {
  const source = context && typeof context === "object" && !Array.isArray(context)
    ? context
    : {};
  return {
    issueStage: normalizeChecklistIssueStage(source.issueStage),
  };
}

function normalizeChecklistApplicability(applicability) {
  if (
    !applicability ||
    typeof applicability !== "object" ||
    Array.isArray(applicability)
  ) {
    return null;
  }

  const normalized = {};
  ["all", "any", "none"].forEach((key) => {
    const values = Array.isArray(applicability[key]) ? applicability[key] : [];
    const deduped = [];
    const seen = new Set();
    values.forEach((value) => {
      const token = String(value || "").trim();
      if (!token) return;
      const normalizedToken = token.toLowerCase();
      if (seen.has(normalizedToken)) return;
      seen.add(normalizedToken);
      deduped.push(normalizedToken);
    });
    if (deduped.length) {
      normalized[key] = deduped;
    }
  });

  return Object.keys(normalized).length ? normalized : null;
}

function normalizeChecklistApplicabilityOverrides(rawOverrides) {
  if (
    !rawOverrides ||
    typeof rawOverrides !== "object" ||
    Array.isArray(rawOverrides)
  ) {
    return {};
  }

  const overrides = {};
  Object.entries(rawOverrides).forEach(([itemId, value]) => {
    const normalizedKey = String(itemId || "").trim();
    const normalizedValue = String(value || "").trim().toLowerCase();
    if (!normalizedKey) return;
    if (normalizedValue !== "show" && normalizedValue !== "hide") return;
    overrides[normalizedKey] = normalizedValue;
  });
  return overrides;
}

function normalizeChecklistInstance(instance = {}) {
  const source = instance && typeof instance === "object" && !Array.isArray(instance)
    ? instance
    : {};
  const checklistId = String(source.checklistId || "").trim();
  if (!checklistId) return null;

  const completedItems = [];
  const seenCompleted = new Set();
  (Array.isArray(source.completedItems) ? source.completedItems : []).forEach((itemId) => {
    const normalizedId = String(itemId || "").trim();
    if (!normalizedId || seenCompleted.has(normalizedId)) return;
    seenCompleted.add(normalizedId);
    completedItems.push(normalizedId);
  });

  const itemNotesSource =
    source.itemNotes && typeof source.itemNotes === "object" && !Array.isArray(source.itemNotes)
      ? source.itemNotes
      : {};
  const itemNotes = {};
  Object.entries(itemNotesSource).forEach(([itemId, note]) => {
    const normalizedId = String(itemId || "").trim();
    if (!normalizedId) return;
    itemNotes[normalizedId] = String(note ?? "");
  });

  return {
    instanceId:
      typeof source.instanceId === "string" && source.instanceId.trim()
        ? source.instanceId.trim()
        : generateChecklistInstanceId(),
    checklistId,
    completedItems,
    itemNotes,
    applicabilityOverrides: normalizeChecklistApplicabilityOverrides(
      source.applicabilityOverrides
    ),
  };
}

function normalizeChecklistRowType(type) {
  return String(type || "").trim().toLowerCase() === CHECKLIST_ROW_TYPES.SUBHEADER
    ? CHECKLIST_ROW_TYPES.SUBHEADER
    : CHECKLIST_ROW_TYPES.ITEM;
}

function isChecklistSubheader(item) {
  return normalizeChecklistRowType(item?.type) === CHECKLIST_ROW_TYPES.SUBHEADER;
}

function getChecklistCheckableItems(items = []) {
  return (Array.isArray(items) ? items : []).filter(
    (item) => !isChecklistSubheader(item)
  );
}

function getChecklistItemNumberLookup(items = []) {
  const lookup = new Map();
  let itemNumber = 0;

  (Array.isArray(items) ? items : []).forEach((item) => {
    if (isChecklistSubheader(item)) {
      itemNumber = 0;
      return;
    }

    itemNumber += 1;
    lookup.set(item.id, itemNumber);
  });

  return lookup;
}

function getChecklistSections(items = []) {
  const sections = [];
  let currentSection = { header: null, items: [] };

  (Array.isArray(items) ? items : []).forEach((item) => {
    if (isChecklistSubheader(item)) {
      if (currentSection.header || currentSection.items.length) {
        sections.push(currentSection);
      }
      currentSection = { header: item, items: [] };
      return;
    }

    currentSection.items.push({
      item,
      number: currentSection.items.length + 1,
    });
  });

  if (currentSection.header || currentSection.items.length) {
    sections.push(currentSection);
  }

  return sections;
}

function getChecklistSectionBounds(items = [], index) {
  if (!Array.isArray(items) || !items.length || index < 0 || index >= items.length) {
    return null;
  }

  let start = index;
  if (!isChecklistSubheader(items[index])) {
    start = -1;
    for (let i = index; i >= 0; i -= 1) {
      if (isChecklistSubheader(items[i])) {
        start = i;
        break;
      }
    }
    if (start < 0) start = 0;
  }

  let end = items.length - 1;
  for (let i = start + 1; i < items.length; i += 1) {
    if (isChecklistSubheader(items[i])) {
      end = i - 1;
      break;
    }
  }

  return { start, end };
}

function getChecklistDragDescriptor(checklist, itemId) {
  const items = Array.isArray(checklist?.items) ? checklist.items : [];
  const index = items.findIndex((item) => item.id === itemId);
  if (index < 0) return null;

  return {
    type: isChecklistSubheader(items[index]) ? "subheader" : "item",
    itemId,
    startIndex: index,
    endIndex: index,
  };
}

function clearChecklistDragStyles() {
  document.querySelectorAll(".checklist-item-row").forEach((row) => {
    row.classList.remove("checklist-drop-before", "checklist-drop-after", "checklist-dragging");
    delete row.dataset.dropPosition;
  });
}

function isChecklistRowMenuOpen(checklistId, itemId) {
  return (
    checklistRowMenuState?.checklistId === checklistId &&
    checklistRowMenuState?.itemId === itemId
  );
}

function setChecklistRowMenuState(checklistId, itemId) {
  checklistRowMenuState =
    checklistId && itemId ? { checklistId, itemId } : null;
}

function moveChecklistRows(checklistId, draggedItemId, targetItemId, before = true) {
  const checklist = getChecklistById(checklistId);
  if (!checklist) return false;

  const items = Array.isArray(checklist.items) ? checklist.items : [];
  const dragDescriptor = getChecklistDragDescriptor(checklist, draggedItemId);
  const targetIndex = items.findIndex((item) => item.id === targetItemId);
  if (!dragDescriptor || targetIndex < 0) return false;
  if (targetIndex >= dragDescriptor.startIndex && targetIndex <= dragDescriptor.endIndex) {
    return false;
  }

  const movedItems = items.slice(
    dragDescriptor.startIndex,
    dragDescriptor.endIndex + 1
  );
  const remainingItems = items.filter(
    (_, index) => index < dragDescriptor.startIndex || index > dragDescriptor.endIndex
  );
  const remainingTargetIndex = remainingItems.findIndex((item) => item.id === targetItemId);
  if (remainingTargetIndex < 0) return false;

  const insertIndex = before ? remainingTargetIndex : remainingTargetIndex + 1;
  remainingItems.splice(insertIndex, 0, ...movedItems);

  checklist.items = remainingItems;
  reindexChecklistItems(checklist.items);
  saveChecklists();
  return true;
}

function createChecklistTemplateSnapshot(items = []) {
  return (Array.isArray(items) ? items : []).map((item, order) => {
    const normalizedItem = normalizeChecklistItem(
      typeof item === "string"
        ? item
        : {
            text: item?.text,
            type: item?.type,
            applicability: item?.applicability,
          },
      order
    );
    const snapshot = {
      text: normalizedItem.item.text,
      type: normalizedItem.item.type,
    };
    if (normalizedItem.item.applicability) {
      snapshot.applicability = normalizedItem.item.applicability;
    }
    return snapshot;
  });
}

function createChecklistItemsFromTexts(texts = []) {
  return createChecklistTemplateSnapshot(texts).map((item, order) =>
    normalizeChecklistItem(item, order).item
  );
}

function normalizeChecklistItem(item, order) {
  let changed = false;
  const base =
    item && typeof item === "object" && !Array.isArray(item) ? { ...item } : {};
  if (base !== item) changed = true;
  const rawId = typeof base.id === "string" ? base.id : "";
  const id = rawId.trim() || generateChecklistItemId();
  if (!rawId.trim() || rawId.trim() !== rawId) changed = true;
  const text = typeof item === "string" ? item : base.text;
  if (typeof text !== "string") changed = true;
  const rawType = typeof item === "string" ? CHECKLIST_ROW_TYPES.ITEM : base.type;
  const type = normalizeChecklistRowType(rawType);
  if (type !== rawType) changed = true;
  const normalizedApplicability = normalizeChecklistApplicability(base.applicability);
  if (
    stableStringify(normalizedApplicability) !==
    stableStringify(base.applicability ?? null)
  ) {
    changed = true;
  }
  if (base.order !== order) changed = true;
  const normalizedItem = {
    ...base,
    id,
    text: String(text ?? ""),
    order,
    type,
  };
  if (normalizedApplicability) {
    normalizedItem.applicability = normalizedApplicability;
  } else if (Object.prototype.hasOwnProperty.call(normalizedItem, "applicability")) {
    delete normalizedItem.applicability;
    changed = true;
  }
  return {
    changed,
    item: normalizedItem,
  };
}

function reindexChecklistItems(items = []) {
  items.forEach((item, index) => {
    item.order = index;
  });
  return items;
}

function isLegacyDefaultChecklistName(name) {
  const normalized = String(name || "").trim().toLowerCase();
  return normalized === "" || normalized === "default" || normalized === "default electrical";
}

function normalizeChecklistRecord(rawChecklist, index) {
  let changed = false;
  const source =
    rawChecklist && typeof rawChecklist === "object" && !Array.isArray(rawChecklist)
      ? rawChecklist
      : {};
  if (source !== rawChecklist) changed = true;

  let id = typeof source.id === "string" ? source.id.trim() : "";
  if (!id) {
    id = generateChecklistId();
    changed = true;
  }

  let name = typeof source.name === "string" ? source.name.trim() : "";
  if (!name) {
    name = `Checklist ${index + 1}`;
    changed = true;
  }

  if (!Array.isArray(source.items)) changed = true;
  const items = (Array.isArray(source.items) ? source.items : []).map((item, itemIndex) => {
    const normalizedItem = normalizeChecklistItem(item, itemIndex);
    changed = changed || normalizedItem.changed;
    return normalizedItem.item;
  });

  let templateKey =
    typeof source.templateKey === "string" ? source.templateKey.trim() : "";
  if (templateKey && !getBundledChecklistTemplateByKey(templateKey)) {
    templateKey = "";
    changed = true;
  }

  const isLegacyDefault =
    id === LEGACY_DEFAULT_CHECKLIST_ID || source.isDefault === true;
  if (isLegacyDefault) {
    if (!templateKey) {
      templateKey = CHECKLIST_TEMPLATE_KEYS.PRE_FLIGHT_ELECTRICAL;
      changed = true;
    }
    if (isLegacyDefaultChecklistName(name)) {
      if (name !== PRE_FLIGHT_ELECTRICAL_CHECKLIST_NAME) changed = true;
      name = PRE_FLIGHT_ELECTRICAL_CHECKLIST_NAME;
    }
    if (source.isDefault === true || source.isLocked === true) changed = true;
  }

  return {
    changed,
    checklist: {
      ...source,
      id,
      name,
      templateKey: templateKey || null,
      items,
      isDefault: false,
      isLocked: false,
    },
  };
}

function normalizeChecklistsDb(rawDb) {
  let changed = false;
  const source =
    rawDb && typeof rawDb === "object" && !Array.isArray(rawDb) ? rawDb : {};
  if (source !== rawDb) changed = true;

  const rawChecklists = Array.isArray(source.checklists) ? source.checklists : [];
  if (!Array.isArray(source.checklists)) changed = true;

  const usedIds = new Set();
  const normalizedChecklists = [];
  rawChecklists.forEach((rawChecklist, index) => {
    const normalized = normalizeChecklistRecord(rawChecklist, index);
    changed = changed || normalized.changed;
    let nextChecklist = normalized.checklist;
    if (usedIds.has(nextChecklist.id)) {
      nextChecklist = { ...nextChecklist, id: generateChecklistId() };
      changed = true;
    }
    usedIds.add(nextChecklist.id);
    normalizedChecklists.push(nextChecklist);
  });

  const rawTemplateOverrides =
    source.templateOverrides &&
    typeof source.templateOverrides === "object" &&
    !Array.isArray(source.templateOverrides)
      ? source.templateOverrides
      : {};
  if (rawTemplateOverrides !== source.templateOverrides) changed = true;

  const normalizedTemplateOverrides = {};
  CHECKLIST_TEMPLATES.forEach((template) => {
    if (!Object.prototype.hasOwnProperty.call(rawTemplateOverrides, template.key)) return;
    const normalizedOverride = normalizeChecklistTemplateOverride(
      rawTemplateOverrides[template.key],
      template.key
    );
    changed = changed || normalizedOverride.changed;
    if (normalizedOverride.override) {
      normalizedTemplateOverrides[template.key] = normalizedOverride.override;
    } else {
      changed = true;
    }
  });

  return {
    changed,
    data: {
      checklists: normalizedChecklists,
      templateOverrides: normalizedTemplateOverrides,
      lastModified:
        typeof source.lastModified === "string" ? source.lastModified : null,
    },
  };
}

async function loadChecklists() {
  try {
    const rawData = await window.pywebview.api.get_checklists();
    const normalized = normalizeChecklistsDb(rawData || {});
    checklistsDb = normalized.data;
    if (normalized.changed) {
      await saveChecklists();
    }
    return checklistsDb;
  } catch (e) {
    console.warn("Failed to load checklists:", e);
    checklistsDb = { checklists: [], templateOverrides: {}, lastModified: null };
    return checklistsDb;
  }
}

async function saveChecklists({
  skipCloud = false,
  saveTimestamp = true,
  timestamp = new Date().toISOString(),
  silent = false,
} = {}) {
  const resolvedTimestamp =
    normalizeIsoTimestamp(timestamp) || new Date().toISOString();
  const comparableChanged =
    getCloudComparableFingerprint("checklists") !==
    lastCloudComparableFingerprints.checklists;
  if (saveTimestamp && comparableChanged) {
    touchLocalSyncTimestamp("checklists", resolvedTimestamp);
    checklistsDb.lastModified = resolvedTimestamp;
  }
  try {
    const response = await window.pywebview.api.save_checklists(checklistsDb);
    if (response.status !== "success") throw new Error(response.message);
    syncCloudComparableFingerprint("checklists");
    if (
      !skipCloud &&
      comparableChanged &&
      cloudSyncState.enabled &&
      !isCloudSyncApplying()
    ) {
      queueCloudStatePush("checklists");
    }
    return true;
  } catch (e) {
    console.warn("Failed to save checklists:", e);
    if (!silent) {
      toast("Failed to save checklists.");
    }
    return false;
  }
}

function getChecklistById(id) {
  return checklistsDb.checklists.find((c) => c.id === id);
}

function findChecklistByTemplateKey(templateKey) {
  const normalizedKey = String(templateKey || "").trim();
  if (!normalizedKey) return null;
  return (
    checklistsDb.checklists.find(
      (checklist) => String(checklist?.templateKey || "").trim() === normalizedKey
    ) || null
  );
}

function getOrCreateChecklistByTemplateKey(templateKey) {
  const existing = findChecklistByTemplateKey(templateKey);
  if (existing) return existing;
  return createChecklistFromTemplate(templateKey);
}

function isAciesElectricalChecklist(checklist) {
  return (
    String(checklist?.templateKey || "").trim() ===
    CHECKLIST_TEMPLATE_KEYS.ACIES_ELECTRICAL
  );
}

function ensureProjectChecklistProfiles(project) {
  if (!project) return null;
  project.checklistProfiles = normalizeProjectChecklistProfiles(project.checklistProfiles);
  return project.checklistProfiles;
}

function getProjectAciesElectricalChecklistProfile(project) {
  const profiles = ensureProjectChecklistProfiles(project);
  return profiles?.[ACIES_ELECTRICAL_SCOPE_PROFILE_KEY] || null;
}

function ensureDeliverableChecklistContext(deliverable) {
  if (!deliverable) return null;
  deliverable.checklistContext = normalizeChecklistContext(deliverable.checklistContext);
  return deliverable.checklistContext;
}

function getAppliedChecklistEntryByTemplateKey(deliverable, templateKey) {
  if (!deliverable || !Array.isArray(deliverable.appliedChecklists)) return null;
  for (let instanceIndex = 0; instanceIndex < deliverable.appliedChecklists.length; instanceIndex += 1) {
    const instance = deliverable.appliedChecklists[instanceIndex];
    const checklist = getChecklistById(instance?.checklistId);
    if (!checklist) continue;
    if (String(checklist.templateKey || "").trim() !== String(templateKey || "").trim()) {
      continue;
    }
    return { instance, checklist, instanceIndex };
  }
  return null;
}

function buildChecklistScopeAnswerState(project, deliverable) {
  const projectProfile = getProjectAciesElectricalChecklistProfile(project);
  const answers = normalizeAciesElectricalChecklistAnswers(projectProfile?.answers);
  return {
    ...answers,
    issueStage: normalizeChecklistIssueStage(deliverable?.checklistContext?.issueStage),
  };
}

function buildChecklistScopeTokenSet(answerState = {}) {
  const tokens = new Set();
  Object.entries(answerState).forEach(([fieldKey, rawValue]) => {
    const field = getChecklistScopeFieldConfig(fieldKey);
    if (field && field.tokenizable === false) return;
    const normalizedValue =
      fieldKey === "issueStage"
        ? normalizeChecklistIssueStage(rawValue)
        : normalizeChecklistScopeOptionValue(fieldKey, rawValue);
    if (!normalizedValue) return;
    tokens.add(`${fieldKey}:${normalizedValue}`);
  });
  return tokens;
}

function hasChecklistScopeAnswer(answerState = {}, fieldKey = "") {
  if (fieldKey === "issueStage") {
    return Boolean(normalizeChecklistIssueStage(answerState.issueStage));
  }
  return Boolean(normalizeChecklistScopeFieldValue(fieldKey, answerState[fieldKey]));
}

function getChecklistApplicabilityTokenCategory(token) {
  const normalized = String(token || "").trim().toLowerCase();
  const separatorIndex = normalized.indexOf(":");
  return separatorIndex >= 0 ? normalized.slice(0, separatorIndex) : normalized;
}

function evaluateChecklistApplicabilityClause(mode, tokens, context) {
  const normalizedTokens = (Array.isArray(tokens) ? tokens : [])
    .map((token) => String(token || "").trim().toLowerCase())
    .filter(Boolean);
  if (!normalizedTokens.length) {
    return { status: "pass", reason: "" };
  }

  const answerTokens = context?.answerTokens || new Set();
  const answerState = context?.answerState || {};

  if (mode === "all") {
    let unresolved = false;
    for (const token of normalizedTokens) {
      if (answerTokens.has(token)) continue;
      const category = getChecklistApplicabilityTokenCategory(token);
      if (!hasChecklistScopeAnswer(answerState, category)) {
        unresolved = true;
        continue;
      }
      return { status: "filtered", reason: "Required scope mismatch." };
    }
    return unresolved
      ? { status: "unresolved", reason: "Needs more scope answers." }
      : { status: "pass", reason: "" };
  }

  if (mode === "any") {
    if (normalizedTokens.some((token) => answerTokens.has(token))) {
      return { status: "pass", reason: "" };
    }
    const unresolved = normalizedTokens.some((token) => {
      const category = getChecklistApplicabilityTokenCategory(token);
      return !hasChecklistScopeAnswer(answerState, category);
    });
    return unresolved
      ? { status: "unresolved", reason: "Needs more scope answers." }
      : { status: "filtered", reason: "Not relevant to the selected scope." };
  }

  if (mode === "none") {
    if (normalizedTokens.some((token) => answerTokens.has(token))) {
      return { status: "filtered", reason: "Excluded by the selected scope." };
    }
    const unresolved = normalizedTokens.some((token) => {
      const category = getChecklistApplicabilityTokenCategory(token);
      return !hasChecklistScopeAnswer(answerState, category);
    });
    return unresolved
      ? { status: "unresolved", reason: "Needs more scope answers." }
      : { status: "pass", reason: "" };
  }

  return { status: "pass", reason: "" };
}

function evaluateChecklistItemApplicability(item, context, overrides = {}) {
  if (!item || isChecklistSubheader(item)) {
    return {
      visibility: "applicable",
      manualOverride: "",
      badge: "",
      reason: "",
    };
  }

  const overrideValue =
    overrides && typeof overrides === "object"
      ? normalizeChecklistApplicabilityOverrides(overrides)[item.id]
      : "";
  if (overrideValue === "show") {
    return {
      visibility: "applicable",
      manualOverride: "show",
      badge: "Shown manually",
      reason: "Shown for this deliverable.",
    };
  }
  if (overrideValue === "hide") {
    return {
      visibility: "filtered",
      manualOverride: "hide",
      badge: "Hidden manually",
      reason: "Hidden for this deliverable.",
    };
  }

  const applicability = normalizeChecklistApplicability(item.applicability);
  if (!applicability) {
    return {
      visibility: "applicable",
      manualOverride: "",
      badge: "",
      reason: "",
    };
  }

  const results = [
    evaluateChecklistApplicabilityClause("all", applicability.all, context),
    evaluateChecklistApplicabilityClause("any", applicability.any, context),
    evaluateChecklistApplicabilityClause("none", applicability.none, context),
  ];
  const filteredResult = results.find((result) => result.status === "filtered");
  if (filteredResult) {
    return {
      visibility: "filtered",
      manualOverride: "",
      badge: "Filtered",
      reason: filteredResult.reason,
    };
  }
  const unresolvedResult = results.find((result) => result.status === "unresolved");
  if (unresolvedResult) {
    return {
      visibility: "unresolved",
      manualOverride: "",
      badge: "Needs scope",
      reason: unresolvedResult.reason,
    };
  }

  return {
    visibility: "applicable",
    manualOverride: "",
    badge: "",
    reason: "",
  };
}

function buildWorkroomChecklistViewModel(project, deliverable, instance, checklist) {
  const sections = getChecklistSections(Array.isArray(checklist?.items) ? checklist.items : []);
  const completedSet = new Set(
    Array.isArray(instance?.completedItems) ? instance.completedItems : []
  );
  const hasFiltering = isAciesElectricalChecklist(checklist);
  const answerState = hasFiltering
    ? buildChecklistScopeAnswerState(project, deliverable)
    : null;
  const context = hasFiltering
    ? {
        answerState,
        answerTokens: buildChecklistScopeTokenSet(answerState),
      }
    : null;
  const overrides =
    instance && typeof instance === "object" ? instance.applicabilityOverrides || {} : {};

  const summary = {
    visibleTotal: 0,
    visibleCompletedTotal: 0,
    filteredTotal: 0,
    unresolvedTotal: 0,
    applicableTotal: 0,
  };

  const viewSections = sections.map((section) => {
    const visibleItems = [];
    const filteredItems = [];
    (Array.isArray(section.items) ? section.items : []).forEach((entry) => {
      const evaluation = hasFiltering
        ? evaluateChecklistItemApplicability(entry.item, context, overrides)
        : {
            visibility: "applicable",
            manualOverride: "",
            badge: "",
            reason: "",
          };
      const nextEntry = {
        ...entry,
        evaluation,
        isCompleted: completedSet.has(entry.item.id),
      };
      if (evaluation.visibility === "filtered") {
        filteredItems.push(nextEntry);
        summary.filteredTotal += 1;
        return;
      }
      visibleItems.push(nextEntry);
      summary.visibleTotal += 1;
      if (nextEntry.isCompleted) summary.visibleCompletedTotal += 1;
      if (evaluation.visibility === "unresolved") {
        summary.unresolvedTotal += 1;
      } else {
        summary.applicableTotal += 1;
      }
    });

    return {
      header: section.header,
      visibleItems,
      filteredItems,
    };
  });

  return {
    hasFiltering,
    context,
    sections: viewSections,
    summary,
  };
}

function createChecklist(name, options = {}) {
  const template =
    typeof options.templateKey === "string"
      ? getChecklistTemplateByKey(options.templateKey)
      : null;
  const normalizedItems = (Array.isArray(options.items) ? options.items : []).map(
    (item, index) => normalizeChecklistItem(item, index).item
  );
  const newChecklist = {
    id:
      typeof options.id === "string" && options.id.trim()
        ? options.id.trim()
        : generateChecklistId(),
    name:
      String(name || "").trim() || `Checklist ${checklistsDb.checklists.length + 1}`,
    templateKey: template?.key || null,
    isDefault: false,
    isLocked: false,
    items: normalizedItems,
  };
  checklistsDb.checklists.push(newChecklist);
  if (options.saveNow !== false) {
    saveChecklists();
  }
  return newChecklist;
}

function createChecklistFromTemplate(templateKey) {
  const template = getChecklistTemplateByKey(templateKey);
  if (!template) return null;
  return createChecklist(template.name, {
    templateKey: template.key,
    items: createChecklistItemsFromTexts(template.items),
  });
}

function removeChecklistFromAllDeliverables(checklistId) {
  let changed = false;
  if (!Array.isArray(db) || !checklistId) return changed;

  db.forEach((project) => {
    const deliverables = getProjectDeliverables(project);
    if (!Array.isArray(deliverables)) return;
    deliverables.forEach((deliverable) => {
      if (!Array.isArray(deliverable.appliedChecklists)) return;
      const originalCount = deliverable.appliedChecklists.length;
      deliverable.appliedChecklists = deliverable.appliedChecklists.filter(
        (instance) => instance?.checklistId !== checklistId
      );
      if (deliverable.appliedChecklists.length !== originalCount) {
        changed = true;
      }
    });
  });

  return changed;
}

function deleteChecklist(id) {
  const idx = checklistsDb.checklists.findIndex((c) => c.id === id);
  if (idx > -1) {
    checklistsDb.checklists.splice(idx, 1);
    if (removeChecklistFromAllDeliverables(id)) {
      save();
    }
    saveChecklists();
    return true;
  }
  return false;
}

function addChecklistItem(checklistId, text) {
  return addChecklistRow(checklistId, text, CHECKLIST_ROW_TYPES.ITEM);
}

function addChecklistSubheader(checklistId, text) {
  return addChecklistRow(checklistId, text, CHECKLIST_ROW_TYPES.SUBHEADER);
}

function insertChecklistItem(checklistId, insertIndex, text = "New checklist item") {
  return insertChecklistRow(
    checklistId,
    insertIndex,
    CHECKLIST_ROW_TYPES.ITEM,
    text
  );
}

function insertChecklistRow(
  checklistId,
  insertIndex,
  type = CHECKLIST_ROW_TYPES.ITEM,
  text = ""
) {
  const checklist = getChecklistById(checklistId);
  if (!checklist) return null;

  const items = Array.isArray(checklist.items) ? checklist.items : [];
  const boundedIndex = Math.max(0, Math.min(Number(insertIndex) || 0, items.length));
  const newItem = normalizeChecklistItem(
    {
      text,
      type,
    },
    boundedIndex
  ).item;

  items.splice(boundedIndex, 0, newItem);
  reindexChecklistItems(items);
  saveChecklists();
  return newItem;
}

function insertChecklistSubheader(checklistId, insertIndex, text = "New section") {
  return insertChecklistRow(
    checklistId,
    insertIndex,
    CHECKLIST_ROW_TYPES.SUBHEADER,
    text
  );
}

function addChecklistRow(checklistId, text, type = CHECKLIST_ROW_TYPES.ITEM) {
  const checklist = getChecklistById(checklistId);
  const nextIndex = Array.isArray(checklist?.items) ? checklist.items.length : 0;
  return insertChecklistRow(checklistId, nextIndex, type, text);
}

function removeChecklistItem(checklistId, itemId) {
  const checklist = getChecklistById(checklistId);
  if (checklist) {
    const idx = checklist.items.findIndex(i => i.id === itemId);
    if (idx > -1) {
      checklist.items.splice(idx, 1);
      reindexChecklistItems(checklist.items);
      saveChecklists();
      return true;
    }
  }
  return false;
}

function updateChecklistItem(checklistId, itemId, text) {
  const checklist = getChecklistById(checklistId);
  if (checklist) {
    const item = checklist.items.find((i) => i.id === itemId);
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
  const template = getChecklistTemplateByKey(checklist?.templateKey);
  if (checklist && template) {
    checklist.items = createChecklistItemsFromTexts(template.items);
    saveChecklists();
    return true;
  }
  return false;
}

function overrideChecklistTemplate(checklistId) {
  const checklist = getChecklistById(checklistId);
  const templateKey =
    typeof checklist?.templateKey === "string" ? checklist.templateKey.trim() : "";
  const baseTemplate = getBundledChecklistTemplateByKey(templateKey);
  if (!checklist || !baseTemplate) return false;

  if (
    !checklistsDb.templateOverrides ||
    typeof checklistsDb.templateOverrides !== "object" ||
    Array.isArray(checklistsDb.templateOverrides)
  ) {
    checklistsDb.templateOverrides = {};
  }

  checklistsDb.templateOverrides[templateKey] = {
    name: String(checklist.name || "").trim() || baseTemplate.name,
    items: createChecklistTemplateSnapshot(checklist.items),
    updatedAt: new Date().toISOString(),
  };

  saveChecklists();
  return true;
}

// Timesheet Constants
const DISCIPLINE_TO_FUNCTION = {
  Electrical: "E",
  Mechanical: "M",
  Plumbing: "P",
};
const WORKROOM_PHASES = WORKROOM_PHASE_OPTIONS.reduce((lookup, phase, index) => {
  lookup[phase.value] = {
    ...phase,
    index,
  };
  return lookup;
}, {});
const WORKROOM_PRE_DESIGN_PAGES = [
  {
    key: "scope_jurisdiction",
    title: "Scope & Jurisdiction",
    description:
      "Capture the core project type, jurisdiction, and utility context before design work begins.",
  },
  {
    key: "building_profile",
    title: "Building Profile",
    description:
      "Record the building scale and service characteristics that shape the electrical approach.",
  },
  {
    key: "existing_conditions",
    title: "Existing Conditions",
    description:
      "Identify the major existing systems and scope flags that affect the design checklist.",
  },
];
const WORKROOM_CAD_DISCIPLINES = ["Electrical", "Mechanical", "Plumbing"];
const WORKROOM_AUTO_SELECT_CAD_TOOL_IDS = new Set([
  "toolPublishDwgs",
  "toolFreezeLayers",
  "toolThawLayers",
]);
const WORKROOM_CAD_TOOL_IDS = new Set([
  ...WORKROOM_AUTO_SELECT_CAD_TOOL_IDS,
  "toolCleanXrefs",
]);
const WORKROOM_TEMPLATE_TOOL_IDS = new Set([
  "toolCreateNarrativeTemplate",
  "toolCreatePlanCheckTemplate",
]);
const WORKROOM_LAUNCH_CONTEXT_TOOL_IDS = new Set([
  ...WORKROOM_CAD_TOOL_IDS,
  ...WORKROOM_TEMPLATE_TOOL_IDS,
  "toolBackupDrawings",
]);
const WORKROOM_HIDDEN_TOOL_IDS = new Set([]);
const WORKROOM_PHASE_CHECKLIST_MAP = {
  pre_design: null,
  design: {
    templateKey: CHECKLIST_TEMPLATE_KEYS.ELECTRICAL_GENERAL,
    title: "Electrical TI - General Checklist",
    intro:
      "Use this checklist to work the project from layout through coordination while work items stay docked to the right.",
  },
  preflight: {
    templateKey: CHECKLIST_TEMPLATE_KEYS.PRE_FLIGHT_ELECTRICAL,
    title: "Preflight Electrical Checklist",
    intro:
      "Run this final issue pass before publishing so the permit set is coordinated, complete, and ready to go out.",
  },
  post_permit: null,
};
const WORKROOM_PHASE_TOOL_MAP = {
  pre_design: ["toolCleanXrefs"],
  design: ["toolLightingSchedule", "toolCircuitBreaker"],
  preflight: ["toolFreezeLayers", "toolPublishDwgs", "toolThawLayers"],
  post_permit: ["toolCreateNarrativeTemplate", "toolCreatePlanCheckTemplate"],
};
const WORKROOM_ALWAYS_AVAILABLE_TOOLS = [
  "toolCopyProjectLocally",
  "toolBackupDrawings",
  "toolWireSizer",
  "toolCircuitBreaker",
];
const MAX_HOURS_PER_DAY = 24;
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

function getWorkroomConfiguredDisciplines() {
  const resolved = [];
  const seen = new Set();
  normalizeDisciplineList(userSettings.discipline).forEach((discipline) => {
    const normalized = normalizeWorkroomCadDiscipline(discipline, "");
    if (!normalized || !WORKROOM_CAD_DISCIPLINES.includes(normalized)) return;
    if (seen.has(normalized)) return;
    seen.add(normalized);
    resolved.push(normalized);
  });
  return resolved;
}

function getWorkroomAvailableDisciplines() {
  const configured = getWorkroomConfiguredDisciplines();
  if (!configured.length) return [...WORKROOM_CAD_DISCIPLINES];
  if (configured.length === 1) {
    const primary = configured[0];
    return [
      primary,
      ...WORKROOM_CAD_DISCIPLINES.filter((discipline) => discipline !== primary),
    ];
  }
  return configured;
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

function formatTimesheetProjectName(project) {
  const name = String(project?.name || "").trim();
  const nick = String(project?.nick || "").trim();
  if (name && nick && name.toLowerCase() !== nick.toLowerCase()) {
    return `${name} (${nick})`;
  }
  return name || nick || "";
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
  if (!projectName) return "";

  const trailingNickname = projectName.match(/\(([^()]+)\)\s*$/);
  if (trailingNickname) {
    const nick = String(trailingNickname[1] || "").trim();
    if (nick) return `name:${nick.toLowerCase()}`;
  }

  return `name:${projectName.toLowerCase()}`;
}

function getDuplicateTimesheetProjectIds(entries) {
  const projectIds = new Map();

  (entries || []).forEach((entry) => {
    const projectId = String(entry?.projectId || "").trim();
    if (!projectId) return;

    const key = projectId.toLowerCase();
    if (!projectIds.has(key)) {
      projectIds.set(key, { displayValue: projectId, count: 0 });
    }

    projectIds.get(key).count += 1;
  });

  return Array.from(projectIds.values())
    .filter((item) => item.count > 1)
    .map((item) => item.displayValue)
    .sort((a, b) =>
      a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" })
    );
}

function updateTimesheetsDuplicateIndicator(entries) {
  const timesheetsTabBtn = document.querySelector(
    '.main-tab-btn[data-tab="timesheets"]'
  );
  if (!timesheetsTabBtn) return;

  const duplicateIds = getDuplicateTimesheetProjectIds(entries);
  const hasDuplicates = duplicateIds.length > 0;
  timesheetsTabBtn.classList.toggle(
    "has-duplicate-project-id",
    hasDuplicates
  );

  if (!hasDuplicates) {
    timesheetsTabBtn.removeAttribute("title");
    timesheetsTabBtn.removeAttribute("aria-label");
    return;
  }

  const duplicateLabel = duplicateIds.join(", ");
  const duplicateText =
    duplicateIds.length === 1 ? "Duplicate project ID" : "Duplicate project IDs";

  timesheetsTabBtn.title = `${duplicateText} in current timesheet: ${duplicateLabel}`;
  timesheetsTabBtn.setAttribute(
    "aria-label",
    `Timesheets. Warning: ${duplicateText.toLowerCase()} ${duplicateLabel}`
  );
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

async function saveTimesheets({
  skipCloud = false,
  saveTimestamp = true,
  timestamp = new Date().toISOString(),
  silent = false,
} = {}) {
  const resolvedTimestamp =
    normalizeIsoTimestamp(timestamp) || new Date().toISOString();
  const comparableChanged =
    getCloudComparableFingerprint("timesheets") !==
    lastCloudComparableFingerprints.timesheets;
  if (saveTimestamp && comparableChanged) {
    touchLocalSyncTimestamp("timesheets", resolvedTimestamp);
    timesheetDb.lastModified = resolvedTimestamp;
  }
  try {
    const response = await window.pywebview.api.save_timesheets(timesheetDb);
    if (response.status !== "success") throw new Error(response.message);
    syncCloudComparableFingerprint("timesheets");
    if (
      !skipCloud &&
      comparableChanged &&
      cloudSyncState.enabled &&
      !isCloudSyncApplying()
    ) {
      queueCloudStatePush("timesheets");
    }
    return true;
  } catch (e) {
    console.warn("Failed to save timesheets:", e);
    if (!silent) {
      toast("Failed to save timesheet data.");
    }
    return false;
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

async function saveTemplates({
  skipCloud = false,
  saveTimestamp = true,
  timestamp = new Date().toISOString(),
  silent = false,
} = {}) {
  const resolvedTimestamp =
    normalizeIsoTimestamp(timestamp) || new Date().toISOString();
  const comparableChanged =
    getCloudComparableFingerprint("templates") !==
    lastCloudComparableFingerprints.templates;
  if (saveTimestamp && comparableChanged) {
    touchLocalSyncTimestamp("templates", resolvedTimestamp);
    templatesDb.lastModified = resolvedTimestamp;
  }
  try {
    const response = await window.pywebview.api.save_templates(templatesDb);
    if (response.status !== "success") throw new Error(response.message);
    syncCloudComparableFingerprint("templates");
    if (
      !skipCloud &&
      comparableChanged &&
      cloudSyncState.enabled &&
      !isCloudSyncApplying()
    ) {
      queueCloudStatePush("templates");
    }
    return true;
  } catch (e) {
    console.warn("Failed to save templates:", e);
    if (!silent) {
      toast("Failed to save templates.");
    }
    return false;
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
  const templateNeedsRelink = template.cloudNeedsRelink && !template.sourcePath;

  const copyBtn = el('button', {
    className: 'btn btn-primary tiny',
    textContent: 'Copy to Folder',
    disabled: templateNeedsRelink,
    title: templateNeedsRelink
      ? 'This template synced without a local source file. Re-add or relink it on this device.'
      : 'Copy template to folder',
    onclick: (e) => {
      e.stopPropagation();
      if (templateNeedsRelink) return;
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
}

function normalizeWorkroomFolderPath(rawPath) {
  return normalizeWindowsPath(rawPath);
}

function findWorkroomProjectRootById(projectPath) {
  return findProjectRootPath(projectPath);
}

function getWorkroomRootFolderPath(project) {
  const overridePath = normalizeWorkroomFolderPath(project?.workroomRootPath || "");
  if (overridePath) return overridePath;
  const projectPath = normalizeProjectPath(project?.path || "");
  return findWorkroomProjectRootById(projectPath) || projectPath;
}

function getWorkroomServerProjectPath(project) {
  const projectPath = normalizeProjectPath(project?.path || "");
  if (!projectPath) return "";
  return findWorkroomProjectRootById(projectPath) || projectPath;
}

function getActiveWorkroomContext() {
  const hasProjectIndex =
    activeChecklistProject !== null && activeChecklistProject !== undefined;
  const projectIndex = hasProjectIndex ? Number(activeChecklistProject) : -1;
  const project =
    Number.isInteger(projectIndex) && projectIndex >= 0 ? db[projectIndex] || null : null;
  if (!project) return { project: null, deliverables: [], deliverable: null };

  const deliverables = getProjectDeliverables(project);
  const hasDeliverableIndex =
    activeChecklistDeliverable !== null && activeChecklistDeliverable !== undefined;
  const deliverableIndex = hasDeliverableIndex ? Number(activeChecklistDeliverable) : -1;
  const deliverable =
    Number.isInteger(deliverableIndex) && deliverableIndex >= 0
      ? deliverables[deliverableIndex] || null
      : null;

  return { project, deliverables, deliverable };
}

function ensureWorkroomCadDiscipline(deliverable, disciplines = []) {
  const options = disciplines.length ? disciplines : getWorkroomAvailableDisciplines();
  if (!options.length) return "Electrical";
  if (!deliverable) return options[0];

  const existing = normalizeWorkroomCadDiscipline(deliverable.workroomCadDiscipline, "");
  if (existing && options.includes(existing)) return existing;

  const fallbackDiscipline = options[0];
  if (deliverable.workroomCadDiscipline !== fallbackDiscipline) {
    deliverable.workroomCadDiscipline = fallbackDiscipline;
    debouncedSave();
  }
  return fallbackDiscipline;
}

function buildWorkroomCadLaunchContext() {
  const { project, deliverable } = getActiveWorkroomContext();
  const disciplines = getWorkroomAvailableDisciplines();
  const activeDiscipline = ensureWorkroomCadDiscipline(deliverable, disciplines);
  const projectPath = normalizeWorkroomFolderPath(project?.path || "");
  const rootProjectPath = getWorkroomRootFolderPath(project);

  return {
    source: "workroom",
    projectPath,
    rootProjectPath,
    discipline: activeDiscipline || disciplines[0] || "Electrical",
    cadFilePaths: [],
    projectName: String(project?.name || project?.nick || project?.id || "").trim(),
    deliverableName: String(deliverable?.name || "").trim(),
  };
}

const SHARED_TOOL_LAUNCH_REGISTRY = Object.freeze([
  {
    id: "toolCopyProjectLocally",
    label: "Local Project Manager",
    menuLabel: "Local Projects",
    launchType: "project-manager",
    isReady: true,
  },
  {
    id: "toolPublishDwgs",
    label: "Publish CAD DWGs in Headless Mode",
    menuLabel: "Publish",
    launchType: "user-selects-files",
    isReady: true,
  },
  {
    id: "toolFreezeLayers",
    label: "Freeze Layers in CAD DWGs Headless Mode",
    menuLabel: "Freeze",
    launchType: "user-selects-files",
    isReady: true,
  },
  {
    id: "toolThawLayers",
    label: "Thaw Layers in CAD DWGs Headless Mode",
    menuLabel: "Thaw",
    launchType: "user-selects-files",
    isReady: true,
  },
  {
    id: "toolCleanXrefs",
    label: "Prepare CAD DWG for Reference",
    menuLabel: "Prepare XREFs",
    launchType: "user-selects-files",
    isReady: true,
  },
  {
    id: "toolCreateNarrativeTemplate",
    label: "Create Narrative of Changes Template",
    menuLabel: "Create NOC",
    launchType: "user-selects-folder",
    isReady: true,
  },
  {
    id: "toolCreatePlanCheckTemplate",
    label: "Create Plan Check Comments Template",
    menuLabel: "Create PCC",
    launchType: "user-selects-folder",
    isReady: true,
  },
  {
    id: "toolWireSizer",
    label: "Wire Sizer",
    menuLabel: "Wire Sizer",
    launchType: "modal",
    isReady: true,
  },
  {
    id: "toolCircuitBreaker",
    label: "Panel Schedule AI",
    menuLabel: "Panel Schedule AI",
    launchType: "modal",
    isReady: true,
  },
  {
    id: "toolBackupDrawings",
    label: "Backup Drawings",
    menuLabel: "Backup DWGs",
    launchType: "archive-project",
    isReady: true,
  },
  {
    id: "toolLightingSchedule",
    label: "Lighting Schedule AI",
    menuLabel: "Lighting Schedule AI",
    launchType: "modal",
    isReady: false,
  },
  {
    id: "toolTitle24Compliance",
    label: "Title 24 Compliance",
    menuLabel: "Title 24 Compliance",
    launchType: "modal",
    isReady: false,
  },
]);

function getSharedToolLaunchEntry(toolId) {
  return (
    SHARED_TOOL_LAUNCH_REGISTRY.find(
      (entry) => String(entry?.id || "").trim() === String(toolId || "").trim()
    ) || null
  );
}

function getReadySharedToolLaunchEntries() {
  return SHARED_TOOL_LAUNCH_REGISTRY.filter((entry) => entry.isReady === true);
}

function getDeliverableToolMenuEntries() {
  const deliverableMenuOrder = [
    "toolPublishDwgs",
    "toolFreezeLayers",
    "toolThawLayers",
    "toolCleanXrefs",
    "toolCreateNarrativeTemplate",
    "toolCreatePlanCheckTemplate",
    "toolWireSizer",
    "toolCircuitBreaker",
    "toolBackupDrawings",
    "toolCopyProjectLocally",
  ];
  const readyEntryMap = new Map(
    getReadySharedToolLaunchEntries()
      .filter((entry) => entry.includeInDeliverableMenu !== false)
      .map((entry) => [String(entry.id || "").trim(), entry])
  );
  return deliverableMenuOrder
    .map((toolId) => readyEntryMap.get(toolId) || null)
    .filter(Boolean);
}

function getLaunchContextProjectRoot(launchContext = null) {
  if (!launchContext || typeof launchContext !== "object") return "";
  const source = String(launchContext?.source || "").trim().toLowerCase();
  const rawPath = String(
    launchContext?.rootProjectPath || launchContext?.projectPath || ""
  ).trim();
  if (!rawPath) return "";
  if (source === "workroom") {
    return normalizeWorkroomFolderPath(rawPath);
  }
  return normalizeProjectPath(rawPath) || normalizeWindowsPath(rawPath);
}

function hasLaunchContextProjectPath(launchContext = null) {
  return Boolean(getLaunchContextProjectRoot(launchContext));
}

function buildProjectsTabToolLaunchContext(project, deliverable) {
  const projectPath = normalizeProjectPath(project?.path || "");
  return {
    source: "projects-tab",
    projectPath,
    rootProjectPath: projectPath,
    cadFilePaths: [],
    projectName: String(project?.name || project?.nick || project?.id || "").trim(),
    deliverableName: String(deliverable?.name || "").trim(),
  };
}

function queuePendingCadLaunchContext(launchContext = null) {
  pendingCadLaunchContext = launchContext ? deepCloneJson(launchContext, null) : null;
}

function launchSharedToolCard(toolId, launchContext = null) {
  const entry = getSharedToolLaunchEntry(toolId);
  if (!entry || entry.isReady !== true) return false;
  const card = document.getElementById(entry.id);
  if (!card) {
    toast(`${entry.label} is unavailable.`);
    return false;
  }
  if (launchContext) {
    queuePendingCadLaunchContext(launchContext);
  }
  card.click();
  return true;
}

function consumePendingCadLaunchContext() {
  const context = pendingCadLaunchContext;
  pendingCadLaunchContext = null;
  return context;
}

function resolveCadLaunchContextForTool() {
  const pendingContext = consumePendingCadLaunchContext();
  if (pendingContext) return pendingContext;

  const checklistModal = document.getElementById("checklistModal");
  if (!checklistModal?.open) return null;

  const { project, deliverable } = getActiveWorkroomContext();
  if (!project || !deliverable) return null;

  return buildWorkroomCadLaunchContext();
}

async function setWorkroomLocalProjectPath(project, nextPath, { saveNow = true } = {}) {
  if (!project) return false;
  const normalizedNext = normalizeWorkroomFolderPath(nextPath);
  const normalizedCurrent = normalizeWorkroomFolderPath(project.localProjectPath || "");
  if (!normalizedNext || normalizedCurrent === normalizedNext) return false;
  project.localProjectPath = normalizedNext;
  if (saveNow) await save();
  else debouncedSave();
  return true;
}

function renderWorkroomProjectHeader() {}

function findProjectByNormalizedPath(rawPath) {
  const normalizedPath = normalizeProjectPath(rawPath);
  if (!normalizedPath) return null;
  return (
    db.find(
      (project) => normalizeProjectPath(project?.path || "") === normalizedPath
    ) || null
  );
}

function getTemplateToolContext(options = {}) {
  const launchContext = options?.launchContext || null;
  const launchSource = String(launchContext?.source || "").trim().toLowerCase();
  const launchProjectPath = getLaunchContextProjectRoot(launchContext);
  if (launchProjectPath) {
    const project =
      findProjectByNormalizedPath(launchContext?.projectPath || "") ||
      findProjectByNormalizedPath(launchContext?.rootProjectPath || "") ||
      findProjectByNormalizedPath(launchProjectPath);
    return {
      projectPath:
        launchProjectPath ||
        (launchSource === "workroom" ? getWorkroomRootFolderPath(project) : ""),
      projectName: String(
        launchContext?.projectName ||
          project?.name ||
          project?.nick ||
          project?.id ||
          ""
      ).trim(),
      deliverableName: String(launchContext?.deliverableName || "").trim(),
      source: launchSource || "default",
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
  const activityId = beginActivity({
    activityId: createActivityId(toolId || `template_${templateKey}`),
    toolId,
    label,
    message: "Initializing...",
    progress: 5,
  });
  if (card) {
    card.classList.add("running");
  }

  try {
    updateActivity(activityId, {
      message: "Loading templates...",
      progress: 12,
    });
    await loadTemplates();
  } catch (e) {
    console.warn("Failed to refresh templates:", e);
  }

  const template = getTemplateByKey(templateKey);
  if (!template) {
    failActivity(activityId, {
      message: `Template "${label}" not found.`,
    });
    return;
  }

  const { projectPath, projectName, deliverableName } = getTemplateToolContext({
    launchContext,
  });
  const context = {};
  if (projectName) context.projectName = projectName;
  if (deliverableName) context.deliverableName = deliverableName;
  const defaultFolder = projectPath || null;
  let selection = null;
  try {
    updateActivity(activityId, {
      message: "Select output folder...",
      progress: 18,
    });
    selection = await window.pywebview.api.select_template_output_folder(
      defaultFolder
    );
  } catch (e) {
    failActivity(activityId, {
      message: "Error selecting output folder.",
    });
    return;
  }

  if (!selection || selection.status === "cancelled") {
    acceptActivity(activityId);
    clearTemplateToolRunState(toolId);
    return;
  }
  if (selection.status !== "success" || !selection.path) {
    failActivity(activityId, {
      message: selection.message || "Failed to select output folder.",
    });
    return;
  }

  const templateOptions = { templateKey };

  try {
    updateActivity(activityId, {
      message: "Saving template...",
      progress: 45,
    });
    const result = await window.pywebview.api.copy_template_to_folder(
      template.id,
      selection.path,
      null,
      context,
      templateOptions
    );
    if (result.status === "success") {
      const warnings = Array.isArray(result.warnings)
        ? result.warnings.filter(Boolean)
        : [];
      const openFolderPath = String(
        result.outputFolderPath || result.openedFolderPath || selection.path || ""
      ).trim();
      if (warnings.length) {
        completeActivity(activityId, {
          status: ACTIVITY_STATUS.WARNING,
          message: `Template saved. ${warnings.join(" ")}`,
          openFolderPath,
        });
      } else {
        completeActivity(activityId, {
          message: "Template saved.",
          openFolderPath,
        });
      }
      return result;
    } else {
      failActivity(activityId, {
        message: result.message || "Failed to save template.",
      });
    }
  } catch (e) {
    failActivity(activityId, {
      message: "Error saving template.",
    });
  } finally {
    clearTemplateToolRunState(toolId);
  }
}

// ===================== TIMESHEET RENDERING =====================

function renderTimesheets() {
  const weekKey = formatWeekKey(currentTimesheetWeek);
  const entries = getWeekEntries(weekKey);
  updateTimesheetsDuplicateIndicator(entries);

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
    updateTimesheetsDuplicateIndicator(
      getWeekEntries(formatWeekKey(currentTimesheetWeek))
    );
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
    const hours = normalizeTimesheetHours(entry.hours?.[day] || 0);
    cell.appendChild(createHourInput(entry, day, hours, row));
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

function normalizeTimesheetHours(value) {
  const parsed = Number.parseFloat(value);
  if (!Number.isFinite(parsed)) return 0;
  return Math.min(Math.max(Math.round(parsed * 10) / 10, 0), MAX_HOURS_PER_DAY);
}

function formatTimesheetHours(value) {
  const normalized = normalizeTimesheetHours(value);
  if (normalized <= 0) return "";
  return Number.isInteger(normalized) ? String(normalized) : normalized.toFixed(1);
}

function getRemainingTimesheetHoursForDay(day, currentEntry) {
  const weekKey = formatWeekKey(currentTimesheetWeek);
  const entries = getWeekEntries(weekKey);
  const otherHours = entries.reduce((sum, entry) => {
    if (entry === currentEntry || entry?.id === currentEntry?.id) return sum;
    return sum + normalizeTimesheetHours(entry.hours?.[day] || 0);
  }, 0);
  return Math.max(0, MAX_HOURS_PER_DAY - otherHours);
}

function refreshTimesheetRowTotal(row, entry) {
  const rowTotalCell = row?.querySelector(".ts-row-total");
  if (!rowTotalCell) return;
  const total = calculateRowTotal(entry);
  rowTotalCell.textContent = total > 0 ? total.toFixed(1) : "";
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

function createHourInput(entry, day, hours, row) {
  const input = el("input", {
    className: "ts-hour-input",
    type: "number",
    min: "0",
    max: String(MAX_HOURS_PER_DAY),
    step: "0.1",
    value: formatTimesheetHours(hours),
    placeholder: "0",
    inputMode: "decimal",
    "aria-label": `${day.toUpperCase()} hours`,
  });

  input.oninput = (e) => {
    const rawValue = e.target.value;
    if (!entry.hours) entry.hours = {};

    if (rawValue === "") {
      entry.hours[day] = 0;
      refreshTimesheetRowTotal(row, entry);
      updateTimesheetTotals(getWeekEntries(formatWeekKey(currentTimesheetWeek)));
      saveTimesheets();
      return;
    }

    const parsedHours = Number.parseFloat(rawValue);
    if (!Number.isFinite(parsedHours)) return;

    const allowedHours = getRemainingTimesheetHoursForDay(day, entry);
    const newHours = Math.min(normalizeTimesheetHours(parsedHours), allowedHours);
    entry.hours[day] = newHours;
    refreshTimesheetRowTotal(row, entry);
    updateTimesheetTotals(getWeekEntries(formatWeekKey(currentTimesheetWeek)));
    saveTimesheets();
  };

  input.onblur = (e) => {
    e.target.value = formatTimesheetHours(entry.hours?.[day] || 0);
  };

  return input;
}

function calculateRowTotal(entry) {
  if (!entry.hours) return 0;
  return Object.values(entry.hours).reduce((sum, h) => sum + (h || 0), 0);
}

function updateTimesheetTotals(entries) {
  const dayOrder = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"];
  const totals = { week: 0 };

  dayOrder.forEach((day) => (totals[day] = 0));

  entries.forEach((entry) => {
    dayOrder.forEach((day) => {
      totals[day] += entry.hours?.[day] || 0;
    });
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

  const projectIdsWithEntries = new Set();
  const summaryEntriesByProject = new Map();
  entries.forEach((entry) => {
    const projectIdKey = String(entry?.projectId || "")
      .trim()
      .toLowerCase();
    if (projectIdKey) projectIdsWithEntries.add(projectIdKey);
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
    const entryStatus = getSummaryEntryStatus(summaryEntries, deliverableNames);
    const projectIdKey = String(group.project?.id || "")
      .trim()
      .toLowerCase();
    const hasProjectEntry = projectIdKey
      ? projectIdsWithEntries.has(projectIdKey)
      : entryStatus.count > 0;
    const card = createSuggestionCard(
      group.project,
      deliverableNames,
      group.earliestDue,
      { ...entryStatus, hasProjectEntry }
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
    hasProjectEntry = false,
  } = entryStatus;
  const showAddedStyle = hasProjectEntry;
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
        textContent: "Click to add another entry",
      })
    );
  } else if (hasProjectEntry) {
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
  const projectName = formatTimesheetProjectName(project);

  const referenceEntry = summaryEntries[0];

  entries.push(
    buildProjectSummaryEntry({
      projectId: referenceEntry?.projectId || projectId,
      projectName: projectName || referenceEntry?.projectName || "",
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
  const formattedProjectName = formatTimesheetProjectName(project);
  const projectName = formattedProjectName || referenceEntry?.projectName || "";
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
  const totals = {};
  dayOrder.forEach((day) => (totals[day] = 0));

  entries.forEach((entry) => {
    dayOrder.forEach((day) => (totals[day] += entry.hours?.[day] || 0));
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
const EXPENSE_IMAGE_THUMB_MAX_SIZE = 320;
const EXPENSE_IMAGE_MODAL_MAX_SIZE = 1800;
const expenseAttachmentPreviewCache = new Map();
const expenseAttachmentResolvedPathCache = new Map();
let activeExpenseImagePreviewRequestId = 0;
const expenseImagePreviewState = {
  scale: 1,
  minScale: 1,
  maxScale: 6,
  translateX: 0,
  translateY: 0,
  naturalWidth: 0,
  naturalHeight: 0,
  baseWidth: 0,
  baseHeight: 0,
  stageWidth: 0,
  stageHeight: 0,
  isDragging: false,
  pointerId: null,
  dragStartX: 0,
  dragStartY: 0,
  dragOriginX: 0,
  dragOriginY: 0,
  initialized: false,
};

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
  const previewButton = el("button", {
    className: "expense-image-preview-btn",
    type: "button",
    title: getExpenseAttachmentFilename(image),
    "aria-label": `Preview ${getExpenseAttachmentFilename(image)}`
  });
  thumb.appendChild(previewButton);

  if (isExpensePreviewableImage(image)) {
    setExpenseImageThumbLoadingState(previewButton);
    previewButton.onclick = async () => {
      await openExpenseImagePreview(image);
    };
    void hydrateExpenseImageThumb(previewButton, image);
  } else {
    renderExpenseAttachmentPlaceholder(previewButton, "PDF", "pdf");
    previewButton.onclick = async () => {
      await openExpenseAttachment(image);
    };
  }

  const removeBtn = el("button", {
    className: "remove-image-btn",
    type: "button",
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

function getExpenseAttachmentFilename(image) {
  const explicitName = String(image?.filename || "").trim();
  if (explicitName) return explicitName;
  const rawPath = String(image?.path || "").trim();
  if (!rawPath) return "Attachment";
  return rawPath.split(/[\\/]/).pop() || "Attachment";
}

function getExpenseAttachmentExtension(image) {
  const filename = getExpenseAttachmentFilename(image).toLowerCase();
  const dotIndex = filename.lastIndexOf(".");
  return dotIndex >= 0 ? filename.slice(dotIndex + 1) : "";
}

function isExpensePreviewableImage(image) {
  const ext = getExpenseAttachmentExtension(image);
  return !!String(image?.path || "").trim() && ext !== "pdf";
}

function getExpenseAttachmentPreviewCacheKey(path, maxSize) {
  return `${maxSize}:${String(path || "").trim()}`;
}

function cacheResolvedExpenseAttachmentPath(rawPath, resolvedPath) {
  const normalizedRawPath = String(rawPath || "").trim();
  const normalizedResolvedPath = String(resolvedPath || "").trim();
  if (!normalizedRawPath || !normalizedResolvedPath) return;
  expenseAttachmentResolvedPathCache.set(normalizedRawPath, normalizedResolvedPath);
}

async function fetchExpenseAttachmentPreview(path, maxSize = EXPENSE_IMAGE_THUMB_MAX_SIZE) {
  const rawPath = String(path || "").trim();
  if (!rawPath) {
    return { status: "error", message: "Attachment path is missing." };
  }

  const cacheKey = getExpenseAttachmentPreviewCacheKey(rawPath, maxSize);
  const cached = expenseAttachmentPreviewCache.get(cacheKey);
  if (cached) {
    return cached instanceof Promise ? await cached : cached;
  }

  const request = (async () => {
    if (!window.pywebview?.api?.get_expense_image_preview) {
      return { status: "error", message: "Preview API unavailable." };
    }
    try {
      const result = await window.pywebview.api.get_expense_image_preview(rawPath, maxSize);
      if (result?.path) {
        cacheResolvedExpenseAttachmentPath(rawPath, result.path);
      }
      return result || { status: "error", message: "Preview unavailable." };
    } catch (error) {
      return { status: "error", message: error?.message || "Preview unavailable." };
    }
  })();

  expenseAttachmentPreviewCache.set(cacheKey, request);
  const result = await request;
  expenseAttachmentPreviewCache.set(cacheKey, result);
  return result;
}

async function resolveExpenseAttachmentPath(path) {
  const rawPath = String(path || "").trim();
  if (!rawPath) return "";

  const cached = expenseAttachmentResolvedPathCache.get(rawPath);
  if (cached) return cached;

  if (!window.pywebview?.api?.resolve_expense_attachment_path) {
    return rawPath;
  }

  try {
    const result = await window.pywebview.api.resolve_expense_attachment_path(rawPath);
    if (result?.status === "success" && result.path) {
      cacheResolvedExpenseAttachmentPath(rawPath, result.path);
      return result.path;
    }
  } catch (error) {
    console.error("Error resolving expense attachment path:", error);
  }

  return rawPath;
}

function setExpenseImageThumbLoadingState(previewButton) {
  previewButton.disabled = false;
  previewButton.classList.add("is-loading");
  previewButton.classList.remove("is-error");
  previewButton.replaceChildren(
    el("div", { className: "expense-attachment-placeholder" }, [
      el("span", { className: "expense-attachment-placeholder-icon", textContent: "IMG" }),
      el("span", { className: "expense-attachment-placeholder-label", textContent: "Loading" }),
    ])
  );
}

function renderExpenseAttachmentPlaceholder(previewButton, label, tone = "") {
  previewButton.classList.remove("is-loading");
  previewButton.classList.toggle("is-error", tone === "error");
  previewButton.replaceChildren(
    el("div", {
      className: `expense-attachment-placeholder${tone ? ` ${tone}` : ""}`
    }, [
      el("span", {
        className: "expense-attachment-placeholder-icon",
        textContent: tone === "pdf" ? "PDF" : "IMG"
      }),
      el("span", {
        className: "expense-attachment-placeholder-label",
        textContent: label
      }),
    ])
  );
}

async function hydrateExpenseImageThumb(previewButton, image) {
  const filename = getExpenseAttachmentFilename(image);
  const rawPath = String(image?.path || "").trim();
  try {
    const preview = await fetchExpenseAttachmentPreview(rawPath, EXPENSE_IMAGE_THUMB_MAX_SIZE);
    if (!previewButton.isConnected) return;
    previewButton.classList.remove("is-loading");
    if (preview?.path) {
      image.resolvedPath = preview.path;
      cacheResolvedExpenseAttachmentPath(rawPath, preview.path);
    }
    if (preview?.status === "success" && preview.dataUrl) {
      previewButton.classList.remove("is-error");
      previewButton.replaceChildren(
        el("img", {
          src: preview.dataUrl,
          alt: filename,
          loading: "lazy",
          draggable: false,
        })
      );
      return;
    }

    if (preview?.status === "unsupported") {
      renderExpenseAttachmentPlaceholder(previewButton, "Open file", "pdf");
      previewButton.onclick = async () => {
        await openExpenseAttachment(image);
      };
      return;
    }

    renderExpenseAttachmentPlaceholder(previewButton, "Open file", "error");
    previewButton.onclick = async () => {
      await openExpenseAttachment(image);
    };
  } catch (error) {
    console.error("Error loading expense preview:", error);
    if (!previewButton.isConnected) return;
    renderExpenseAttachmentPlaceholder(previewButton, "Open file", "error");
    previewButton.onclick = async () => {
      await openExpenseAttachment(image);
    };
  }
}

async function openExpenseAttachment(image) {
  const rawPath = String(image?.path || "").trim();
  if (!rawPath || !window.pywebview?.api?.open_path) {
    toast("Unable to open attachment.");
    return;
  }

  try {
    const targetPath = await resolveExpenseAttachmentPath(image.resolvedPath || rawPath);
    const result = await window.pywebview.api.open_path(targetPath);
    if (result?.status && result.status !== "success") {
      throw new Error(result.message || "Unable to open attachment.");
    }
  } catch (error) {
    console.error("Error opening expense attachment:", error);
    toast(error?.message || "Unable to open attachment.");
  }
}

function getExpenseImagePreviewElements() {
  return {
    dialog: document.getElementById("expenseImagePreviewDlg"),
    stage: document.getElementById("expenseImagePreviewStage"),
    img: document.getElementById("expenseImagePreviewImg"),
  };
}

function clampExpenseImagePreviewTranslation(nextX, nextY) {
  const displayedWidth =
    expenseImagePreviewState.baseWidth * expenseImagePreviewState.scale;
  const displayedHeight =
    expenseImagePreviewState.baseHeight * expenseImagePreviewState.scale;
  const maxTranslateX = Math.max(
    0,
    (displayedWidth - expenseImagePreviewState.stageWidth) / 2
  );
  const maxTranslateY = Math.max(
    0,
    (displayedHeight - expenseImagePreviewState.stageHeight) / 2
  );

  return {
    translateX: Math.max(-maxTranslateX, Math.min(maxTranslateX, nextX)),
    translateY: Math.max(-maxTranslateY, Math.min(maxTranslateY, nextY)),
  };
}

function renderExpenseImagePreviewTransform() {
  const { stage, img } = getExpenseImagePreviewElements();
  if (!stage || !img) return;

  stage.classList.toggle(
    "is-zoomed",
    expenseImagePreviewState.scale > expenseImagePreviewState.minScale + 0.001
  );
  stage.classList.toggle("is-dragging", expenseImagePreviewState.isDragging);

  if (!img.hidden) {
    img.style.width = `${expenseImagePreviewState.baseWidth}px`;
    img.style.height = `${expenseImagePreviewState.baseHeight}px`;
    img.style.transform = `translate(${expenseImagePreviewState.translateX}px, ${expenseImagePreviewState.translateY}px) scale(${expenseImagePreviewState.scale})`;
  }
}

function syncExpenseImagePreviewLayout() {
  const { stage, img } = getExpenseImagePreviewElements();
  if (
    !stage ||
    !img ||
    img.hidden ||
    !expenseImagePreviewState.naturalWidth ||
    !expenseImagePreviewState.naturalHeight
  ) {
    return;
  }

  expenseImagePreviewState.stageWidth = Math.max(stage.clientWidth, 1);
  expenseImagePreviewState.stageHeight = Math.max(stage.clientHeight, 1);

  const fitScale = Math.min(
    expenseImagePreviewState.stageWidth / expenseImagePreviewState.naturalWidth,
    expenseImagePreviewState.stageHeight / expenseImagePreviewState.naturalHeight
  );

  expenseImagePreviewState.baseWidth = Math.max(
    1,
    expenseImagePreviewState.naturalWidth * fitScale
  );
  expenseImagePreviewState.baseHeight = Math.max(
    1,
    expenseImagePreviewState.naturalHeight * fitScale
  );

  if (expenseImagePreviewState.scale <= expenseImagePreviewState.minScale) {
    expenseImagePreviewState.translateX = 0;
    expenseImagePreviewState.translateY = 0;
  } else {
    const clamped = clampExpenseImagePreviewTranslation(
      expenseImagePreviewState.translateX,
      expenseImagePreviewState.translateY
    );
    expenseImagePreviewState.translateX = clamped.translateX;
    expenseImagePreviewState.translateY = clamped.translateY;
  }

  renderExpenseImagePreviewTransform();
}

function resetExpenseImagePreviewTransform() {
  expenseImagePreviewState.scale = expenseImagePreviewState.minScale;
  expenseImagePreviewState.translateX = 0;
  expenseImagePreviewState.translateY = 0;
  syncExpenseImagePreviewLayout();
}

function applyExpenseImagePreviewPan(deltaX, deltaY) {
  if (expenseImagePreviewState.scale <= expenseImagePreviewState.minScale) {
    expenseImagePreviewState.translateX = 0;
    expenseImagePreviewState.translateY = 0;
    renderExpenseImagePreviewTransform();
    return;
  }

  const clamped = clampExpenseImagePreviewTranslation(
    expenseImagePreviewState.translateX + deltaX,
    expenseImagePreviewState.translateY + deltaY
  );
  expenseImagePreviewState.translateX = clamped.translateX;
  expenseImagePreviewState.translateY = clamped.translateY;
  renderExpenseImagePreviewTransform();
}

function applyExpenseImagePreviewZoom(nextScale, clientX, clientY) {
  const { stage } = getExpenseImagePreviewElements();
  if (!stage || !expenseImagePreviewState.baseWidth || !expenseImagePreviewState.baseHeight) {
    return;
  }

  const boundedScale = Math.max(
    expenseImagePreviewState.minScale,
    Math.min(expenseImagePreviewState.maxScale, nextScale)
  );
  const previousScale = expenseImagePreviewState.scale;
  if (Math.abs(boundedScale - previousScale) < 0.001) return;

  syncExpenseImagePreviewLayout();

  const stageRect = stage.getBoundingClientRect();
  const pointerX =
    Math.max(stageRect.left, Math.min(stageRect.right, clientX)) -
    (stageRect.left + stageRect.width / 2);
  const pointerY =
    Math.max(stageRect.top, Math.min(stageRect.bottom, clientY)) -
    (stageRect.top + stageRect.height / 2);
  const contentX = (pointerX - expenseImagePreviewState.translateX) / previousScale;
  const contentY = (pointerY - expenseImagePreviewState.translateY) / previousScale;

  expenseImagePreviewState.scale = boundedScale;
  expenseImagePreviewState.translateX = pointerX - contentX * boundedScale;
  expenseImagePreviewState.translateY = pointerY - contentY * boundedScale;

  if (boundedScale <= expenseImagePreviewState.minScale) {
    expenseImagePreviewState.translateX = 0;
    expenseImagePreviewState.translateY = 0;
  } else {
    const clamped = clampExpenseImagePreviewTranslation(
      expenseImagePreviewState.translateX,
      expenseImagePreviewState.translateY
    );
    expenseImagePreviewState.translateX = clamped.translateX;
    expenseImagePreviewState.translateY = clamped.translateY;
  }

  renderExpenseImagePreviewTransform();
}

function stopExpenseImagePreviewDrag(pointerId = null) {
  const { stage } = getExpenseImagePreviewElements();
  if (
    pointerId !== null &&
    expenseImagePreviewState.pointerId !== null &&
    pointerId !== expenseImagePreviewState.pointerId
  ) {
    return;
  }

  if (
    stage &&
    expenseImagePreviewState.pointerId !== null &&
    stage.hasPointerCapture?.(expenseImagePreviewState.pointerId)
  ) {
    stage.releasePointerCapture(expenseImagePreviewState.pointerId);
  }

  expenseImagePreviewState.isDragging = false;
  expenseImagePreviewState.pointerId = null;
  renderExpenseImagePreviewTransform();
}

function resetExpenseImagePreviewDialog() {
  const { stage, img } = getExpenseImagePreviewElements();
  activeExpenseImagePreviewRequestId += 1;
  stopExpenseImagePreviewDrag();

  expenseImagePreviewState.scale = expenseImagePreviewState.minScale;
  expenseImagePreviewState.translateX = 0;
  expenseImagePreviewState.translateY = 0;
  expenseImagePreviewState.naturalWidth = 0;
  expenseImagePreviewState.naturalHeight = 0;
  expenseImagePreviewState.baseWidth = 0;
  expenseImagePreviewState.baseHeight = 0;
  expenseImagePreviewState.stageWidth = 0;
  expenseImagePreviewState.stageHeight = 0;

  if (stage) {
    stage.classList.remove("is-zoomed", "is-dragging");
  }

  if (img) {
    img.hidden = true;
    img.removeAttribute("src");
    img.style.removeProperty("width");
    img.style.removeProperty("height");
    img.style.removeProperty("transform");
  }
}

function closeExpenseImagePreviewDialog() {
  const { dialog } = getExpenseImagePreviewElements();
  if (!dialog) return;

  if (dialog.open) {
    dialog.close();
  } else {
    resetExpenseImagePreviewDialog();
  }
}

function openExpenseImagePreviewDialog({ dataUrl, filename, width, height }) {
  const { dialog, img } = getExpenseImagePreviewElements();
  if (!dialog || !img || !dataUrl) return;

  stopExpenseImagePreviewDrag();
  expenseImagePreviewState.naturalWidth = Math.max(1, Number(width) || 1);
  expenseImagePreviewState.naturalHeight = Math.max(1, Number(height) || 1);
  img.alt = filename || "Expense attachment preview";
  img.src = dataUrl;
  img.hidden = false;

  showDialog(dialog);
  requestAnimationFrame(() => {
    resetExpenseImagePreviewTransform();
  });
}

function loadExpenseImagePreviewSource(dataUrl) {
  return new Promise((resolve, reject) => {
    const preload = new Image();
    preload.onload = () => {
      resolve({
        dataUrl,
        width: preload.naturalWidth || preload.width || 1,
        height: preload.naturalHeight || preload.height || 1,
      });
    };
    preload.onerror = () => reject(new Error("Unable to load preview image."));
    preload.src = dataUrl;
  });
}

function initExpenseImagePreviewDialog() {
  if (expenseImagePreviewState.initialized) return;

  const { dialog, stage } = getExpenseImagePreviewElements();
  if (!dialog || !stage) return;

  expenseImagePreviewState.initialized = true;

  dialog.addEventListener("close", () => {
    resetExpenseImagePreviewDialog();
  });

  dialog.addEventListener(
    "wheel",
    (event) => {
      event.preventDefault();
      event.stopPropagation();
    },
    { passive: false }
  );

  dialog.addEventListener("click", (event) => {
    if (event.target === dialog) {
      closeExpenseImagePreviewDialog();
    }
  });

  stage.addEventListener(
    "wheel",
    (event) => {
      event.preventDefault();
      event.stopPropagation();
      const zoomFactor = event.deltaY < 0 ? 1.18 : 1 / 1.18;
      applyExpenseImagePreviewZoom(
        expenseImagePreviewState.scale * zoomFactor,
        event.clientX,
        event.clientY
      );
    },
    { passive: false }
  );

  stage.addEventListener("pointerdown", (event) => {
    if (event.button !== 0 || expenseImagePreviewState.scale <= expenseImagePreviewState.minScale) {
      return;
    }
    event.preventDefault();
    event.stopPropagation();

    expenseImagePreviewState.isDragging = true;
    expenseImagePreviewState.pointerId = event.pointerId;
    expenseImagePreviewState.dragStartX = event.clientX;
    expenseImagePreviewState.dragStartY = event.clientY;
    expenseImagePreviewState.dragOriginX = expenseImagePreviewState.translateX;
    expenseImagePreviewState.dragOriginY = expenseImagePreviewState.translateY;
    stage.setPointerCapture?.(event.pointerId);
    renderExpenseImagePreviewTransform();
  });

  stage.addEventListener("pointermove", (event) => {
    if (
      !expenseImagePreviewState.isDragging ||
      expenseImagePreviewState.pointerId !== event.pointerId
    ) {
      return;
    }

    event.preventDefault();
    expenseImagePreviewState.translateX = expenseImagePreviewState.dragOriginX;
    expenseImagePreviewState.translateY = expenseImagePreviewState.dragOriginY;
    applyExpenseImagePreviewPan(
      event.clientX - expenseImagePreviewState.dragStartX,
      event.clientY - expenseImagePreviewState.dragStartY
    );
  });

  stage.addEventListener("pointerup", (event) => {
    stopExpenseImagePreviewDrag(event.pointerId);
  });
  stage.addEventListener("pointercancel", (event) => {
    stopExpenseImagePreviewDrag(event.pointerId);
  });
  stage.addEventListener("lostpointercapture", (event) => {
    stopExpenseImagePreviewDrag(event.pointerId);
  });

  window.addEventListener("resize", () => {
    if (dialog.open) {
      syncExpenseImagePreviewLayout();
    }
  });
}

async function openExpenseImagePreview(image) {
  const filename = getExpenseAttachmentFilename(image);
  if (!isExpensePreviewableImage(image)) {
    await openExpenseAttachment(image);
    return;
  }

  const requestId = ++activeExpenseImagePreviewRequestId;

  try {
    const preview = await fetchExpenseAttachmentPreview(
      image.path,
      EXPENSE_IMAGE_MODAL_MAX_SIZE
    );

    if (requestId !== activeExpenseImagePreviewRequestId) return;

    if (preview?.path) {
      image.resolvedPath = preview.path;
      cacheResolvedExpenseAttachmentPath(image.path, preview.path);
    }

    if (preview?.status !== "success" || !preview.dataUrl) {
      throw new Error(preview?.message || "Unable to load preview.");
    }

    const loadedPreview = await loadExpenseImagePreviewSource(preview.dataUrl);
    if (requestId !== activeExpenseImagePreviewRequestId) return;

    openExpenseImagePreviewDialog({
      dataUrl: loadedPreview.dataUrl,
      filename,
      width: loadedPreview.width,
      height: loadedPreview.height,
    });
  } catch (error) {
    console.error("Error opening expense image preview:", error);
    if (requestId !== activeExpenseImagePreviewRequestId) return;
    toast(error?.message || "Unable to load preview. Opening file instead.");
    await openExpenseAttachment(image);
  }
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

function createAttachmentTriggerIcon(size = 15) {
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("viewBox", "0 0 24 24");
  svg.setAttribute("width", String(size));
  svg.setAttribute("height", String(size));
  svg.setAttribute("fill", "none");
  svg.setAttribute("stroke", "currentColor");
  svg.setAttribute("stroke-width", "2");
  svg.setAttribute("stroke-linecap", "round");
  svg.setAttribute("stroke-linejoin", "round");
  svg.setAttribute("class", "attachment-trigger-icon");
  svg.setAttribute("aria-hidden", "true");
  svg.setAttribute("focusable", "false");
  const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
  path.setAttribute(
    "d",
    "m18.375 12.739-7.693 7.693a4.5 4.5 0 0 1-6.364-6.364l10.94-10.94a3 3 0 1 1 4.243 4.243l-9.879 9.879a1.5 1.5 0 0 1-2.121-2.121l8.818-8.818"
  );
  svg.appendChild(path);
  return svg;
}

function createSettingsIcon(size = 14) {
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("viewBox", "0 0 24 24");
  svg.setAttribute("width", String(size));
  svg.setAttribute("height", String(size));
  svg.setAttribute("fill", "none");
  svg.setAttribute("stroke", "currentColor");
  svg.setAttribute("stroke-width", "2");
  svg.setAttribute("stroke-linecap", "round");
  svg.setAttribute("stroke-linejoin", "round");
  svg.setAttribute("aria-hidden", "true");
  svg.setAttribute("focusable", "false");

  const circle = document.createElementNS(
    "http://www.w3.org/2000/svg",
    "circle"
  );
  circle.setAttribute("cx", "12");
  circle.setAttribute("cy", "12");
  circle.setAttribute("r", "3");
  svg.appendChild(circle);

  const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
  path.setAttribute(
    "d",
    "M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 1 1-4 0v-.09a1.65 1.65 0 0 0-1-1.51 1.65 1.65 0 0 0-1.82.33l-.06.06A2 2 0 1 1 4.38 17l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 1 1 0-4h.09a1.65 1.65 0 0 0 1.51-1 1.65 1.65 0 0 0-.33-1.82L4.21 7.2a2 2 0 1 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 1 1 4 0v.09a1.65 1.65 0 0 0 1 1.51h.08a1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9c0 .66.39 1.26 1 1.51H21a2 2 0 1 1 0 4h-.09c-.66 0-1.26.39-1.51 1z"
  );
  svg.appendChild(path);

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
  if (
    joined.includes("outlook.office.com/") ||
    joined.includes("outlook.office365.com/") ||
    joined.includes("outlook.live.com/")
  ) {
    return "outlook-url";
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
  const messageId = String(
    value.messageId || value.message_id || value.graphMessageId || ""
  ).trim();
  const internetMessageId = String(
    value.internetMessageId || value.internet_message_id || ""
  ).trim();
  return { raw, url, label, source, savedAt, messageId, internetMessageId };
}

function normalizeEmailRefKey(value) {
  const ref = normalizeEmailRef(value);
  if (!ref) return "";
  const messageId = String(ref.messageId || "").trim().toLowerCase();
  const internetMessageId = String(ref.internetMessageId || "")
    .trim()
    .toLowerCase();
  if (messageId || internetMessageId) {
    return [`msg:${messageId}`, `internet:${internetMessageId}`].join("|");
  }
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
  return normalizeEmailRefKey(left) === normalizeEmailRefKey(right);
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

function isWebAttachmentTarget(value) {
  return /^https?:\/\//i.test(String(value || "").trim());
}

function normalizeAttachmentTarget(target, type = "path") {
  let raw = String(target || "").trim();
  if (!raw) return "";
  if (type === "url") return raw;
  const localPathFromUrl = fileUrlToLocalPath(raw);
  if (localPathFromUrl) {
    raw = localPathFromUrl;
  }
  if (/^[A-Za-z]:[\\/]/.test(raw) || /^\\\\/.test(raw)) {
    return normalizeWindowsPath(raw);
  }
  return raw;
}

function inferAttachmentType(target, preferredType = "") {
  const normalizedPreferred = String(preferredType || "").trim().toLowerCase();
  if (["file", "folder", "path", "url", "email"].includes(normalizedPreferred)) {
    return normalizedPreferred;
  }
  return isWebAttachmentTarget(target) ? "url" : "path";
}

function getAttachmentTypeLabel(type = "path") {
  if (type === "file") return "File";
  if (type === "folder") return "Folder";
  if (type === "url") return "URL";
  if (type === "email") return "Email";
  return "Path";
}

function getAttachmentDescriptionFallback({
  type = "path",
  target = "",
  emailRef = null,
} = {}) {
  if (type === "email") {
    const normalizedEmailRef = normalizeEmailRef(emailRef);
    return (
      String(normalizedEmailRef?.label || "").trim() ||
      basename(getEmailRefLocalPath(normalizedEmailRef)) ||
      basename(normalizedEmailRef?.raw || normalizedEmailRef?.url || "") ||
      "Email"
    );
  }
  if (type === "url") {
    try {
      const parsed = new URL(String(target || "").trim());
      return parsed.hostname || parsed.href || "Link";
    } catch {
      return String(target || "").trim() || "Link";
    }
  }
  return basename(String(target || "").trim()) || String(target || "").trim() || "Attachment";
}

function getAttachmentEntryKey(attachment) {
  const normalized = normalizeAttachmentEntry(attachment);
  if (!normalized) return "";
  if (normalized.type === "email") {
    return `email:${normalizeEmailRefKey(normalized.emailRef)}`;
  }
  return `${normalized.type}:${String(normalized.target || "").trim().toLowerCase()}`;
}

function normalizeAttachmentEntry(entry) {
  if (!entry) return null;
  const source =
    typeof entry === "string" ? { target: entry, description: "" } : entry;
  const normalizedEmailRef = normalizeEmailRef(source.emailRef || source.ref || null);
  const type = inferAttachmentType(
    source.target || source.raw || source.url || "",
    normalizedEmailRef ? "email" : source.type
  );

  if (type === "email") {
    if (!normalizedEmailRef) return null;
    return {
      id: String(source.id || createId("att")).trim() || createId("att"),
      type: "email",
      description:
        String(source.description || source.label || "").trim() ||
        getAttachmentDescriptionFallback({
          type: "email",
          emailRef: normalizedEmailRef,
        }),
      emailRef: normalizedEmailRef,
    };
  }

  const target = normalizeAttachmentTarget(
    source.target || source.raw || source.url || "",
    type
  );
  if (!target) return null;

  return {
    id: String(source.id || createId("att")).trim() || createId("att"),
    type,
    description:
      String(source.description || source.label || "").trim() ||
      getAttachmentDescriptionFallback({ type, target }),
    target,
  };
}

function normalizeAttachments(
  attachments,
  { legacyLinks = [], legacyEmailRefs = [], legacyEmailRef = null } = {}
) {
  const hasExplicitAttachments = Array.isArray(attachments);
  const normalized = [];
  const seen = new Set();

  const appendAttachment = (candidate) => {
    const next = normalizeAttachmentEntry(candidate);
    if (!next) return;
    const key = getAttachmentEntryKey(next);
    if (!key || seen.has(key)) return;
    seen.add(key);
    normalized.push(next);
  };

  if (hasExplicitAttachments) {
    attachments.forEach(appendAttachment);
    return normalized;
  }

  normalizeDeliverableLinks(legacyLinks).forEach((link) => {
    appendAttachment({
      type: isWebAttachmentTarget(link.raw || link.url || "") ? "url" : "path",
      target: link.raw || link.url || "",
      description: String(link.label || "").trim(),
    });
  });

  normalizeEmailRefs(legacyEmailRefs, legacyEmailRef).forEach((emailRef) => {
    appendAttachment({
      type: "email",
      description: String(emailRef?.label || "").trim(),
      emailRef,
    });
  });

  return normalized;
}

function buildLegacyLinksFromAttachments(attachments = []) {
  return normalizeAttachments(attachments)
    .filter((attachment) => attachment.type !== "email")
    .map((attachment) =>
      normalizeDeliverableLinkEntry({
        label: String(attachment.description || "").trim(),
        raw: attachment.target || "",
      })
    )
    .filter(Boolean);
}

function buildLegacyEmailRefsFromAttachments(attachments = []) {
  return normalizeAttachments(attachments)
    .filter((attachment) => attachment.type === "email")
    .map((attachment) => normalizeEmailRef(attachment.emailRef))
    .filter(Boolean);
}

function syncAttachmentOwnerCompatFields(
  owner,
  {
    includeEmailRefs = true,
    legacyLinks = owner?.links || [],
    legacyEmailRefs = owner?.emailRefs || [],
    legacyEmailRef = owner?.emailRef || null,
  } = {}
) {
  if (!owner || typeof owner !== "object") return [];
  const attachments = normalizeAttachments(owner.attachments, {
    legacyLinks,
    legacyEmailRefs,
    legacyEmailRef,
  });
  owner.attachments = attachments;
  owner.links = buildLegacyLinksFromAttachments(attachments);
  if (includeEmailRefs) {
    const emailRefs = buildLegacyEmailRefsFromAttachments(attachments);
    owner.emailRefs = emailRefs;
    owner.emailRef = emailRefs[0] || null;
  }
  return attachments;
}

function syncProjectAttachmentFields(project) {
  return syncAttachmentOwnerCompatFields(project, {
    includeEmailRefs: false,
    legacyLinks: project?.links || [],
  });
}

function syncDeliverableAttachmentFields(deliverable) {
  return syncAttachmentOwnerCompatFields(deliverable, {
    includeEmailRefs: true,
    legacyLinks: normalizeDeliverableLinks(
      deliverable?.links,
      deliverable?.linkPath || ""
    ),
    legacyEmailRefs: deliverable?.emailRefs,
    legacyEmailRef: deliverable?.emailRef,
  });
}

function syncAttachmentOwnerEmailRefs(owner) {
  if (!owner || typeof owner !== "object") return [];
  syncAttachmentOwnerCompatFields(owner, {
    includeEmailRefs: true,
    legacyLinks: owner?.links || [],
    legacyEmailRefs: owner?.emailRefs,
    legacyEmailRef: owner?.emailRef,
  });
  return owner.emailRefs || [];
}

function syncDeliverableEmailRefs(deliverable) {
  syncDeliverableAttachmentFields(deliverable);
  return deliverable?.emailRefs || [];
}

function syncAttachmentOwnerLinks(owner, legacyLinkPath = "") {
  if (!owner || typeof owner !== "object") return [];
  const includeEmailRefs = "emailRefs" in owner || "emailRef" in owner;
  syncAttachmentOwnerCompatFields(owner, {
    includeEmailRefs,
    legacyLinks: normalizeDeliverableLinks(owner.links, legacyLinkPath),
    legacyEmailRefs: owner?.emailRefs,
    legacyEmailRef: owner?.emailRef,
  });
  return owner.links || [];
}

function getAttachmentOwnerKindLabel(kind = "deliverable") {
  if (kind === "project") return "Project";
  if (kind === "task") return "Task";
  if (kind === "note") return "Note";
  return "Deliverable";
}

function getAttachmentOwnerLabel(descriptor = {}) {
  const { kind = "deliverable", owner = null } = descriptor;
  const deliverable =
    descriptor.deliverable || (kind === "deliverable" ? owner : null);
  const ownerText = String(owner?.text || "").trim();
  const projectName = String(descriptor.project?.name || owner?.name || "").trim();
  const deliverableName = String(deliverable?.name || "").trim();
  if (kind === "project") return projectName || getAttachmentOwnerKindLabel(kind);
  return ownerText || deliverableName || projectName || getAttachmentOwnerKindLabel(kind);
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
  const source = String(ref.source || "").trim().toLowerCase();
  if (
    source === "outlook-desktop" &&
    ref.messageId &&
    window.pywebview?.api?.open_outlook_desktop_message
  ) {
    try {
      const res = await window.pywebview.api.open_outlook_desktop_message(
        ref.messageId
      );
      if (res && res.status === "error") {
        throw new Error(res.message || "Unable to open Desktop Outlook email.");
      }
      return true;
    } catch (e) {
      toast("Could not open Desktop Outlook email.");
      return false;
    }
  }
  if (source === "outlook-desktop" && ref.messageId) {
    toast("Desktop Outlook message opening is unavailable in this environment.");
    return false;
  }
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

function buildAttachmentOwnerEmailContext(descriptor = {}, scope = "projects-tab") {
  const { kind = "deliverable", owner = null, project = null } = descriptor;
  const deliverable =
    descriptor.deliverable || (kind === "deliverable" ? owner : null);
  const context = {
    projectId: String(project?.id || val("f_id") || "").trim(),
    projectName: String(project?.name || val("f_name") || "").trim(),
    deliverableId: String(deliverable?.id || "").trim(),
    deliverableName: String(deliverable?.name || "").trim(),
    scope,
    attachmentOwnerKind: kind,
    attachmentOwnerLabel: getAttachmentOwnerLabel({
      kind,
      owner,
      deliverable,
    }),
  };
  const ownerText = String(owner?.text || "").trim();
  if (ownerText) {
    context.attachmentOwnerText = ownerText;
  }
  if (
    kind !== "deliverable" &&
    Number.isInteger(descriptor.index) &&
    descriptor.index >= 0
  ) {
    context.attachmentOwnerIndex = descriptor.index;
  }
  return context;
}

function buildDeliverableEmailContext(deliverable, project, scope = "projects-tab") {
  return {
    ...buildAttachmentOwnerEmailContext(
      {
        kind: "deliverable",
        owner: deliverable,
        deliverable,
        project,
      },
      scope
    ),
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

function normalizeWindowsPath(rawPath) {
  let normalized = String(rawPath || "")
    .trim()
    .replace(/^['"]+|['"]+$/g, "")
    .replace(/\//g, "\\");
  if (!normalized) return "";
  if (/^[A-Za-z]:\\?$/.test(normalized)) {
    return `${normalized.slice(0, 2)}\\`;
  }
  normalized = normalized.replace(/\\+$/, "");
  return normalized;
}

function getWindowsPathLeaf(rawPath) {
  const normalized = normalizeWindowsPath(rawPath);
  if (!normalized) return "";
  const parts = normalized.split("\\");
  return parts[parts.length - 1] || "";
}

function getWindowsPathParent(rawPath) {
  const normalized = normalizeWindowsPath(rawPath);
  if (!normalized) return "";
  if (/^[A-Za-z]:\\$/.test(normalized)) return "";
  if (/^\\\\[^\\]+\\[^\\]+$/.test(normalized)) return "";
  const idx = normalized.lastIndexOf("\\");
  if (idx <= 0) return "";
  return normalized.slice(0, idx);
}

function findProjectRootPath(rawPath) {
  let current = normalizeWindowsPath(rawPath);
  if (!current) return "";
  while (current) {
    const folderName = getWindowsPathLeaf(current);
    if (PROJECT_ROOT_SEGMENT_REGEX.test(folderName)) {
      return current;
    }
    const parent = getWindowsPathParent(current);
    if (!parent || parent.toLowerCase() === current.toLowerCase()) break;
    current = parent;
  }
  return "";
}

function normalizeProjectPath(rawPath) {
  const normalized = normalizeWindowsPath(rawPath);
  if (!normalized) return "";
  return findProjectRootPath(normalized) || normalized;
}

function parseProjectFromPath(rawPath) {
  const rootPath = findProjectRootPath(rawPath);
  if (!rootPath) return null;
  const match = getWindowsPathLeaf(rootPath).match(PROJECT_ROOT_SEGMENT_REGEX);
  if (!match) return null;
  return {
    id: String(match[1] || "").trim(),
    name: String(match[2] || "").trim(),
  };
}

function getProjectBasePath(rawPath) {
  const normalized = normalizeWindowsPath(rawPath);
  if (!normalized) return "";
  const projectRoot = findProjectRootPath(normalized);
  if (projectRoot) return projectRoot.replace(/\\/g, "/");
  const leaf = getWindowsPathLeaf(normalized);
  const basePath = /\.[A-Za-z0-9]{1,5}$/.test(leaf)
    ? getWindowsPathParent(normalized)
    : normalized;
  return basePath ? basePath.replace(/\\/g, "/") : "";
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
  const shouldSetName = (force || !currentName) && !!parsed.name;

  if (shouldSetId) idInput.value = parsed.id;
  if (shouldSetName) nameInput.value = parsed.name;
  return shouldSetId || shouldSetName;
}

function normalizeProjectPathInput(pathInput, { forceProjectFields = false } = {}) {
  if (!pathInput) return "";
  const rawValue = String(pathInput.value || "");
  applyProjectFromPath(rawValue, { force: forceProjectFields });
  const normalized = normalizeProjectPath(rawValue);
  pathInput.value = normalized;
  pathInput.title = normalized;
  return normalized;
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

function createActivityId(prefix = "activity") {
  const normalizedPrefix = String(prefix || "activity")
    .trim()
    .replace(/[^a-z0-9_-]+/gi, "-")
    .replace(/^-+|-+$/g, "") || "activity";
  if (window.crypto?.randomUUID) {
    return `${normalizedPrefix}_${window.crypto.randomUUID()}`;
  }
  return `${normalizedPrefix}_${Date.now().toString(36)}_${Math.random()
    .toString(36)
    .slice(2, 8)}`;
}

function isTerminalActivityStatus(status) {
  return [ACTIVITY_STATUS.SUCCESS, ACTIVITY_STATUS.WARNING, ACTIVITY_STATUS.ERROR].includes(
    String(status || "").trim().toLowerCase()
  );
}

function clampActivityProgress(value, fallback = 0) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return Math.max(0, Math.min(100, Number(fallback) || 0));
  return Math.max(0, Math.min(100, Math.round(numeric)));
}

function getActivityTrayElements() {
  return {
    tray: document.getElementById("activityTray"),
    toggle: document.getElementById("activityTrayToggle"),
    counts: document.getElementById("activityTrayCounts"),
    body: document.getElementById("activityTrayBody"),
    empty: document.getElementById("activityTrayEmpty"),
    list: document.getElementById("activityTrayList"),
  };
}

function getToolActivityLabel(toolId, fallback = "") {
  const normalizedToolId = String(toolId || "").trim();
  if (normalizedToolId && typeof getSharedToolLaunchEntry === "function") {
    const entry = getSharedToolLaunchEntry(normalizedToolId);
    if (entry?.label) return entry.label;
  }
  return String(fallback || normalizedToolId || "Activity").trim();
}

function getActivityById(activityId) {
  return activityTrayState.items.find((item) => item.id === activityId) || null;
}

function setToolCardRunning(toolId, isRunning) {
  const card = document.getElementById(String(toolId || "").trim());
  if (!card) return;
  card.classList.toggle("running", Boolean(isRunning));
}

function bindToolActivity(toolId, activityId) {
  const normalizedToolId = String(toolId || "").trim();
  const normalizedActivityId = String(activityId || "").trim();
  if (!normalizedToolId || !normalizedActivityId) return;
  activeToolActivityIds.set(normalizedToolId, normalizedActivityId);
}

function releaseToolActivity(toolId, activityId = "") {
  const normalizedToolId = String(toolId || "").trim();
  if (!normalizedToolId) return;
  const boundId = activeToolActivityIds.get(normalizedToolId);
  if (!boundId) return;
  if (activityId && boundId !== activityId) return;
  activeToolActivityIds.delete(normalizedToolId);
}

function sortActivityItems(items = []) {
  return [...items].sort((left, right) => {
    const leftTerminal = isTerminalActivityStatus(left?.status);
    const rightTerminal = isTerminalActivityStatus(right?.status);
    if (leftTerminal !== rightTerminal) return leftTerminal ? 1 : -1;
    return Number(right?.updatedAt || 0) - Number(left?.updatedAt || 0);
  });
}

function toggleActivityTrayCollapsed(nextCollapsed, { force = false } = {}) {
  const { tray, toggle, body } = getActivityTrayElements();
  if (!tray || !toggle || !body) return;
  const hasItems = activityTrayState.items.length > 0;
  if (!hasItems && !force) return;
  activityTrayState.collapsed = Boolean(nextCollapsed);
  tray.classList.toggle("is-collapsed", activityTrayState.collapsed);
  toggle.setAttribute("aria-expanded", String(!activityTrayState.collapsed));
  body.hidden = activityTrayState.collapsed;
}

function maybeExpandActivityTray(reason = "update") {
  if (!activityTrayState.items.length) return;
  if (reason === "first" && !activityTrayState.hasAutoExpanded) {
    activityTrayState.hasAutoExpanded = true;
    toggleActivityTrayCollapsed(false, { force: true });
    return;
  }
  if (reason === "error") {
    toggleActivityTrayCollapsed(false, { force: true });
  }
}

function renderActivityTray() {
  const { tray, counts, empty, list } = getActivityTrayElements();
  if (!tray || !counts || !empty || !list) return;

  const items = sortActivityItems(activityTrayState.items);
  const activeCount = items.filter((item) => !isTerminalActivityStatus(item.status)).length;
  const completedCount = items.length - activeCount;
  counts.textContent = `${activeCount} active, ${completedCount} completed`;

  if (!items.length) {
    tray.hidden = true;
    empty.hidden = false;
    list.replaceChildren();
    toggleActivityTrayCollapsed(false, { force: true });
    activityTrayState.hasAutoExpanded = false;
    return;
  }

  tray.hidden = false;
  empty.hidden = true;
  list.replaceChildren();

  items.forEach((item) => {
    const status = String(item.status || ACTIVITY_STATUS.RUNNING).trim().toLowerCase();
    const iconText =
      status === ACTIVITY_STATUS.SUCCESS
        ? "✓"
        : status === ACTIVITY_STATUS.WARNING
          ? "!"
          : status === ACTIVITY_STATUS.ERROR
            ? "×"
            : "•";

    const card = el("article", {
      className: "activity-card",
      "data-status": status,
    });
    const icon = el("div", {
      className: `activity-card-icon ${status}`,
      textContent: iconText,
      title: status,
    });
    const content = el("div", { className: "activity-card-content" });
    const header = el("div", { className: "activity-card-header" });
    header.append(
      el("div", {
        className: "activity-card-title",
        textContent: item.label || "Activity",
      }),
      el("div", {
        className: "activity-card-percent",
        textContent: `${clampActivityProgress(item.progress, 0)}%`,
      })
    );
    const message = el("div", {
      className: "activity-card-message",
      textContent: item.message || "Working...",
    });
    const progress = el("div", { className: "activity-card-progress" }, [
      el("div", {
        className: "activity-card-progress-bar",
        style: `width:${clampActivityProgress(item.progress, 0)}%`,
      }),
    ]);
    const actions = el("div", { className: "activity-card-actions" });
    if (isTerminalActivityStatus(status) && item.openFolderPath) {
      actions.appendChild(
        el("button", {
          className: "activity-card-action",
          type: "button",
          textContent: item.openFolderLabel || "Open Folder",
          "data-activity-action": "open",
          "data-activity-id": item.id,
        })
      );
    }
    if (isTerminalActivityStatus(status)) {
      actions.appendChild(
        el("button", {
          className: "activity-card-action accept",
          type: "button",
          textContent: "Accept",
          "data-activity-action": "accept",
          "data-activity-id": item.id,
        })
      );
    }

    content.append(header, message, progress);
    if (actions.childNodes.length) {
      content.appendChild(actions);
    }
    card.append(icon, content);
    list.appendChild(card);
  });

  toggleActivityTrayCollapsed(activityTrayState.collapsed, { force: true });
}

function initActivityTray() {
  if (activityTrayState.initialized) return;
  activityTrayState.initialized = true;
  const { toggle, list } = getActivityTrayElements();
  toggle?.addEventListener("click", () => {
    toggleActivityTrayCollapsed(!activityTrayState.collapsed, { force: true });
  });
  list?.addEventListener("click", async (event) => {
    const button = event.target.closest("[data-activity-action]");
    if (!button) return;
    const activityId = String(button.dataset.activityId || "").trim();
    const activity = getActivityById(activityId);
    if (!activity) return;
    const action = String(button.dataset.activityAction || "").trim().toLowerCase();
    if (action === "accept") {
      acceptActivity(activityId);
      return;
    }
    if (action === "open") {
      if (!activity.openFolderPath || !window.pywebview?.api?.open_path) return;
      try {
        await window.pywebview.api.open_path(activity.openFolderPath);
      } catch (error) {
        toast(error?.message || "Unable to open folder.");
      }
    }
  });
  renderActivityTray();
}

function upsertActivity(nextItem, { autoExpandReason = "update" } = {}) {
  const incoming = nextItem && typeof nextItem === "object" ? nextItem : null;
  if (!incoming?.id) return null;

  const now = Date.now();
  const existing = getActivityById(incoming.id);
  const merged = {
    id: incoming.id,
    kind: existing?.kind || incoming.kind || "tool",
    toolId: String(incoming.toolId || existing?.toolId || "").trim(),
    label: String(incoming.label || existing?.label || "Activity").trim(),
    message: String(
      incoming.message == null ? existing?.message || "" : incoming.message
    ).trim(),
    status: String(incoming.status || existing?.status || ACTIVITY_STATUS.RUNNING)
      .trim()
      .toLowerCase(),
    progress: clampActivityProgress(
      incoming.progress == null ? existing?.progress ?? 0 : incoming.progress,
      existing?.progress ?? 0
    ),
    openFolderPath: String(
      incoming.openFolderPath || existing?.openFolderPath || ""
    ).trim(),
    openFolderLabel: String(
      incoming.openFolderLabel || existing?.openFolderLabel || "Open Folder"
    ).trim(),
    panelCount: Math.max(
      Number(incoming.panelCount ?? existing?.panelCount ?? 0) || 0,
      0
    ),
    completedCount: Math.max(
      Number(incoming.completedCount ?? existing?.completedCount ?? 0) || 0,
      0
    ),
    createdAt: Number(existing?.createdAt || now),
    updatedAt: now,
  };

  if (existing) {
    const index = activityTrayState.items.findIndex((item) => item.id === merged.id);
    if (index >= 0) {
      activityTrayState.items.splice(index, 1, merged);
    }
  } else {
    activityTrayState.items.push(merged);
  }

  if (merged.toolId) {
    bindToolActivity(merged.toolId, merged.id);
    setToolCardRunning(merged.toolId, !isTerminalActivityStatus(merged.status));
    if (isTerminalActivityStatus(merged.status)) {
      releaseToolActivity(merged.toolId, merged.id);
    }
  }

  renderActivityTray();
  maybeExpandActivityTray(autoExpandReason);
  return merged;
}

function beginActivity({
  activityId = "",
  toolId = "",
  label = "",
  message = "Starting...",
  progress = 5,
  kind = "tool",
  openFolderPath = "",
  openFolderLabel = "Open Folder",
} = {}) {
  const resolvedId = String(activityId || createActivityId(toolId || kind)).trim();
  return upsertActivity(
    {
      id: resolvedId,
      kind,
      toolId,
      label: getToolActivityLabel(toolId, label),
      message,
      status: ACTIVITY_STATUS.RUNNING,
      progress,
      openFolderPath,
      openFolderLabel,
    },
    { autoExpandReason: activityTrayState.items.length ? "update" : "first" }
  )?.id || resolvedId;
}

function updateActivity(activityId, patch = {}) {
  const existing = getActivityById(activityId);
  if (!existing) return null;
  const nextStatus = String(patch.status || existing.status || ACTIVITY_STATUS.RUNNING)
    .trim()
    .toLowerCase();
  return upsertActivity(
    {
      ...existing,
      ...patch,
      id: existing.id,
      status: nextStatus,
    },
    {
      autoExpandReason:
        nextStatus === ACTIVITY_STATUS.ERROR ? "error" : "update",
    }
  );
}

function completeActivity(activityId, patch = {}) {
  const existing = getActivityById(activityId);
  if (!existing) return null;
  const nextStatus = String(patch.status || ACTIVITY_STATUS.SUCCESS).trim().toLowerCase();
  return upsertActivity(
    {
      ...existing,
      ...patch,
      id: existing.id,
      status: nextStatus,
      progress: 100,
    },
    {
      autoExpandReason: nextStatus === ACTIVITY_STATUS.ERROR ? "error" : "update",
    }
  );
}

function failActivity(activityId, patch = {}) {
  return completeActivity(activityId, {
    ...patch,
    status: ACTIVITY_STATUS.ERROR,
    progress: 100,
  });
}

function acceptActivity(activityId) {
  const existing = getActivityById(activityId);
  if (!existing) return;
  if (existing.toolId) {
    releaseToolActivity(existing.toolId, activityId);
    setToolCardRunning(existing.toolId, false);
  }
  activityTrayState.items = activityTrayState.items.filter((item) => item.id !== activityId);
  renderActivityTray();
}

function getActivityIdForTool(toolId, { create = false, label = "" } = {}) {
  const normalizedToolId = String(toolId || "").trim();
  if (!normalizedToolId) return "";
  const boundId = activeToolActivityIds.get(normalizedToolId);
  if (boundId && getActivityById(boundId)) return boundId;
  if (!create) return "";
  return beginActivity({
    toolId: normalizedToolId,
    label,
    message: "Starting...",
    progress: 5,
  });
}

function normalizeActivityStatusFromPayload(payload = {}, message = "", fallback = ACTIVITY_STATUS.RUNNING) {
  const rawStatus = String(payload?.status || "").trim().toLowerCase();
  if (
    [ACTIVITY_STATUS.RUNNING, ACTIVITY_STATUS.SUCCESS, ACTIVITY_STATUS.WARNING, ACTIVITY_STATUS.ERROR].includes(
      rawStatus
    )
  ) {
    return rawStatus;
  }
  if (message === "DONE") return ACTIVITY_STATUS.SUCCESS;
  if (message.startsWith("WARN:") || message.startsWith("WARNING:")) {
    return ACTIVITY_STATUS.WARNING;
  }
  if (message.startsWith("ERROR:")) return ACTIVITY_STATUS.ERROR;
  return String(fallback || ACTIVITY_STATUS.RUNNING).trim().toLowerCase();
}

function deriveToolActivityProgress(toolId, message, currentProgress = 5) {
  const normalizedToolId = String(toolId || "").trim();
  const text = String(message || "").trim();
  if (!text) return clampActivityProgress(currentProgress, 5);
  if (text === "DONE" || text.startsWith("WARN:") || text.startsWith("ERROR:")) {
    return 100;
  }

  const fileMatch = text.match(/(?:Plotting|Processing)\s+(\d+)\s+of\s+(\d+)\s*:/i);
  if (fileMatch) {
    const completed = Number(fileMatch[1]);
    const total = Math.max(Number(fileMatch[2]), 1);
    if (Number.isFinite(completed) && Number.isFinite(total)) {
      return clampActivityProgress(20 + (completed / total) * 65, currentProgress);
    }
  }

  if (normalizedToolId === "toolPublishDwgs") {
    if (/Waiting for paper size/i.test(text)) return 15;
    if (/Using auto-selected|Received auto-selected|Found AutoCAD|Using specified AutoCAD/i.test(text)) return 10;
    if (/Combining/i.test(text)) return 90;
    if (/Shrinking/i.test(text)) return 95;
    if (/No PDFs were generated/i.test(text)) return 95;
  }

  if (
    ["toolFreezeLayers", "toolThawLayers", "toolCleanXrefs"].includes(normalizedToolId)
  ) {
    if (/Waiting for layer selection|Waiting for user input/i.test(text)) return 15;
    if (/Reading extracted data|Using \d+ DWG|ZIP source selected|Found \d+ DWG/i.test(text)) {
      return 20;
    }
    if (/Successfully processed|Processing \d+ file\(s\)/i.test(text)) return 92;
  }

  if (/Select project folder|Select output folder/i.test(text)) return 10;
  if (/Resolving project folder|Resolving/i.test(text)) return 18;
  if (/Copying|Syncing|Creating archive backup|Saving|Installing|Updating|Uninstalling/i.test(text)) {
    return Math.max(clampActivityProgress(currentProgress, 25), 35);
  }

  return clampActivityProgress(currentProgress, 12);
}

function updateActivityStatusFromPayload(payload = {}) {
  const toolId = String(payload?.toolId || "").trim();
  const activityId =
    String(payload?.activityId || "").trim() ||
    (toolId ? getActivityIdForTool(toolId, { create: true }) : beginActivity({ kind: "activity" }));
  const existing = getActivityById(activityId);
  const rawMessage = String(payload?.message || "").trim();
  const status = normalizeActivityStatusFromPayload(payload, rawMessage, existing?.status);

  let nextMessage = rawMessage;
  let openFolderPath = String(
    payload?.openFolderPath ||
      payload?.outputFolderPath ||
      payload?.bundlePath ||
      payload?.pluginsFolderPath ||
      payload?.archivePath ||
      payload?.localProjectPath ||
      payload?.serverProjectPath ||
      payload?.outputFolder ||
      payload?.folderPath ||
      existing?.openFolderPath ||
      ""
  ).trim();

  if (rawMessage.startsWith("OUTPUT_FOLDER:")) {
    openFolderPath = rawMessage.substring("OUTPUT_FOLDER:".length).trim();
    nextMessage = existing?.message || "Working...";
  } else if (rawMessage.startsWith("WARN:")) {
    nextMessage = rawMessage.substring(5).trim() || "Completed with warnings.";
  } else if (rawMessage.startsWith("WARNING:")) {
    nextMessage = rawMessage.substring(8).trim() || "Completed with warnings.";
  } else if (rawMessage.startsWith("ERROR:")) {
    nextMessage = rawMessage.substring(6).trim() || "Activity failed.";
  } else if (rawMessage === "DONE") {
    nextMessage = payload?.completedMessage || existing?.message || "Done.";
  }

  const explicitProgress = Number(payload?.progress);
  let nextProgress = Number.isFinite(explicitProgress)
    ? clampActivityProgress(explicitProgress, existing?.progress ?? 0)
    : deriveToolActivityProgress(toolId, rawMessage, existing?.progress ?? 5);

  if (toolId === "toolCircuitBreaker") {
    const panelCount = Math.max(Number(payload?.panelCount || existing?.panelCount || 0), 0);
    const completedCount = Math.max(
      Number(
        payload?.completedCount ??
          (Number(payload?.successCount || 0) + Number(payload?.failureCount || 0))
      ) || 0,
      0
    );
    if (status === ACTIVITY_STATUS.RUNNING && panelCount > 0) {
      nextProgress = clampActivityProgress(20 + (completedCount / panelCount) * 70, nextProgress);
    }
  }

  if (!existing) {
    beginActivity({
      activityId,
      toolId,
      label: getToolActivityLabel(toolId, payload?.label || ""),
      message: nextMessage || "Starting...",
      progress: nextProgress,
      openFolderPath,
      openFolderLabel: payload?.openFolderLabel || "Open Folder",
      kind: payload?.kind || "tool",
    });
  }

  const commonPatch = {
    label: getToolActivityLabel(toolId, payload?.label || existing?.label || ""),
    message:
      nextMessage ||
      existing?.message ||
      (status === ACTIVITY_STATUS.ERROR ? "Activity failed." : "Working..."),
    progress: isTerminalActivityStatus(status) ? 100 : nextProgress,
    openFolderPath,
    openFolderLabel: String(
      payload?.openFolderLabel || existing?.openFolderLabel || "Open Folder"
    ).trim(),
    panelCount: payload?.panelCount,
    completedCount: payload?.completedCount,
  };

  if (status === ACTIVITY_STATUS.ERROR) {
    failActivity(activityId, commonPatch);
    return activityId;
  }
  if (status === ACTIVITY_STATUS.SUCCESS || status === ACTIVITY_STATUS.WARNING) {
    completeActivity(activityId, {
      ...commonPatch,
      status,
    });
    return activityId;
  }
  updateActivity(activityId, {
    ...commonPatch,
    status: ACTIVITY_STATUS.RUNNING,
  });
  return activityId;
}

window.updateActivityStatus = function (payload) {
  initActivityTray();
  return updateActivityStatusFromPayload(payload);
};

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
let currentSort = { key: "due", dir: "asc" };
let pinnedProjectDragState = null;
let projectPinHandleSuppressClickUntil = 0;
let statusFilter = "all";
let dueFilter = "all";
let pendingCadLaunchContext = null;
let modalEmailSession = {
  active: false,
  created: new Map(),
  deleteOnSave: new Map(),
};
function createDefaultLocalProjectManagerSyncState() {
  return {
    loadingProjects: false,
    projectsLoaded: false,
    projectsError: "",
    localRootPath: "",
    projectEntries: [],
    searchQuery: "",
    selectedLocalProjectPath: "",
    localProjectPath: "",
    localProjectName: "",
    manualLocalProjectPath: "",
    matchedProjectIndex: -1,
    matchedProjectId: "",
    matchedProjectName: "",
    matchedProjectNick: "",
    manualServerProjectPath: "",
    resolvedServerProjectPath: "",
    serverPathSource: "",
    previewLoading: false,
    previewLoaded: false,
    previewError: "",
    candidateFiles: [],
    blockedEntries: [],
  };
}
const DEFAULT_COPY_PROJECT_LOCALLY_DIALOG_STATE = {
  serverProjectPath: "",
  localProjectPath: "",
  projectName: "",
  resolvedFromWorkroom: false,
  resolutionMode: "",
  workroomProjectPath: "",
  launchContext: null,
  activeTab: "copy",
  copyPreviewLoaded: false,
  copyLoadErrorMessage: "",
  localProjectExists: false,
  folderOptions: [],
  sync: createDefaultLocalProjectManagerSyncState(),
};
let copyProjectLocallyDialogState = {
  ...DEFAULT_COPY_PROJECT_LOCALLY_DIALOG_STATE,
  sync: createDefaultLocalProjectManagerSyncState(),
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
const DEFAULT_CLOUD_SYNC_SETTINGS = {
  enabled: false,
  firebaseUid: "",
  lastSyncedAt: "",
  migrationCompleted: false,
};
const DEFAULT_CLOUD_SYNC_STATE = {
  available: false,
  configured: false,
  enabled: false,
  busy: false,
  status: "local-only",
  message: "Local-only",
  error: "",
  lastSyncedAt: "",
  firebaseUid: "",
  sdkLoaded: false,
  signedIn: false,
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
  separateDeliverableCompletionGroups: true,
  groupDeliverablesByProject: false,
  defaultPmInitials: "",
  cleanDwgOptions: { ...DEFAULT_CLEAN_DWG_OPTIONS },
  publishDwgOptions: { ...DEFAULT_PUBLISH_DWG_OPTIONS },
  freezeLayerOptions: { ...DEFAULT_FREEZE_LAYER_OPTIONS },
  thawLayerOptions: { ...DEFAULT_THAW_LAYER_OPTIONS },
  workroomAutoSelectCadFiles: true,
  enableUnderConstructionTools: false,
  googleAuth: null,
  cloudSync: { ...DEFAULT_CLOUD_SYNC_SETTINGS },
};
const DEFAULT_GOOGLE_AUTH_STATE = {
  signedIn: false,
  provider: "google",
  email: "",
  displayName: "",
  avatarUrl: "",
  signedInAt: "",
  expiresAt: "",
  hasRefreshToken: false,
};
const DEFAULT_OUTLOOK_SCAN_CAPABILITY = {
  loaded: false,
  loading: false,
  desktopAvailable: false,
  desktopReason: "",
};
function formatLocalDateInputValue(date = new Date()) {
  const safeDate = date instanceof Date && !isNaN(date) ? date : new Date();
  const year = safeDate.getFullYear();
  const month = String(safeDate.getMonth() + 1).padStart(2, "0");
  const day = String(safeDate.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function getTodayLocalDateInputValue() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return formatLocalDateInputValue(today);
}

function normalizeOutlookScanDateInput(value = "") {
  const todayValue = getTodayLocalDateInputValue();
  const raw = String(value || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(raw)) return todayValue;
  const [year, month, day] = raw.split("-").map((token) => Number(token));
  const parsed = new Date(year, month - 1, day);
  if (
    !(parsed instanceof Date) ||
    isNaN(parsed) ||
    parsed.getFullYear() !== year ||
    parsed.getMonth() !== month - 1 ||
    parsed.getDate() !== day
  ) {
    return todayValue;
  }
  const normalized = formatLocalDateInputValue(parsed);
  return normalized > todayValue ? todayValue : normalized;
}

function parseOutlookScanDateValue(value = "") {
  const normalized = normalizeOutlookScanDateInput(value);
  const [year, month, day] = normalized.split("-").map((token) => Number(token));
  return new Date(year, month - 1, day, 0, 0, 0, 0);
}

function createEmptyOutlookScanProgress() {
  return {
    active: false,
    stage: "",
    message: "",
    source: "",
    timeframe: "",
    scanDate: "",
    totalEmails: null,
    processedEmails: null,
    includedEmails: null,
    skippedEmails: null,
    deliverablesInPeriod: null,
    relevantEmails: null,
    threadsDetected: null,
    dedupedEmailCount: null,
    dedupeSkippedEmailCount: null,
    receivedAt: "",
  };
}

function normalizeOutlookScanCount(value) {
  if (value === null || value === undefined || value === "") return null;
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return null;
  return Math.max(0, Math.trunc(numeric));
}

function normalizeOutlookScanProgress(raw = {}, previous = null) {
  const base =
    previous && typeof previous === "object"
      ? previous
      : createEmptyOutlookScanProgress();
  const stage =
    raw?.stage !== undefined
      ? String(raw.stage || "").trim().toLowerCase()
      : String(base.stage || "").trim().toLowerCase();
  const message =
    raw?.message !== undefined
      ? String(raw.message || "").trim()
      : String(base.message || "").trim();
  const source =
    raw?.source !== undefined
      ? String(raw.source || "").trim()
      : String(base.source || "").trim();
  const scanDate =
    raw?.scanDate !== undefined
      ? normalizeOutlookScanDateInput(raw.scanDate || "")
      : String(base.scanDate || "").trim();
  const timeframe =
    raw?.timeframe !== undefined
      ? String(raw.timeframe || "").trim().toLowerCase()
      : String(base.timeframe || "").trim().toLowerCase();
  const active =
    typeof raw?.active === "boolean"
      ? raw.active
      : stage
        ? !["done", "error"].includes(stage)
        : !!base.active;
  const receivedAt =
    raw?.receivedAt !== undefined
      ? String(raw.receivedAt || "").trim()
      : String(base.receivedAt || "").trim();
  return {
    ...createEmptyOutlookScanProgress(),
    ...base,
    active,
    stage,
    message,
    source,
    timeframe,
    scanDate,
    totalEmails:
      raw?.totalEmails !== undefined
        ? normalizeOutlookScanCount(raw.totalEmails)
        : normalizeOutlookScanCount(base.totalEmails),
    processedEmails:
      raw?.processedEmails !== undefined
        ? normalizeOutlookScanCount(raw.processedEmails)
        : normalizeOutlookScanCount(base.processedEmails),
    includedEmails:
      raw?.includedEmails !== undefined
        ? normalizeOutlookScanCount(raw.includedEmails)
        : normalizeOutlookScanCount(base.includedEmails),
    skippedEmails:
      raw?.skippedEmails !== undefined
        ? normalizeOutlookScanCount(raw.skippedEmails)
        : normalizeOutlookScanCount(base.skippedEmails),
    deliverablesInPeriod:
      raw?.deliverablesInPeriod !== undefined
        ? normalizeOutlookScanCount(raw.deliverablesInPeriod)
        : normalizeOutlookScanCount(base.deliverablesInPeriod),
    relevantEmails:
      raw?.relevantEmails !== undefined
        ? normalizeOutlookScanCount(raw.relevantEmails)
        : normalizeOutlookScanCount(base.relevantEmails),
    threadsDetected:
      raw?.threadsDetected !== undefined
        ? normalizeOutlookScanCount(raw.threadsDetected)
        : normalizeOutlookScanCount(base.threadsDetected),
    dedupedEmailCount:
      raw?.dedupedEmailCount !== undefined
        ? normalizeOutlookScanCount(raw.dedupedEmailCount)
        : normalizeOutlookScanCount(base.dedupedEmailCount),
    dedupeSkippedEmailCount:
      raw?.dedupeSkippedEmailCount !== undefined
        ? normalizeOutlookScanCount(raw.dedupeSkippedEmailCount)
        : normalizeOutlookScanCount(base.dedupeSkippedEmailCount),
    receivedAt,
  };
}

const DEFAULT_OUTLOOK_SCAN_STATE = {
  mode: "paste",
  scanDate: getTodayLocalDateInputValue(),
  busy: false,
  lastResult: null,
  suggestions: [],
  skipped: [],
  dismissedKeys: [],
  progress: createEmptyOutlookScanProgress(),
  progressLog: [],
  reportOpen: false,
  reportStats: null,
};
let googleAuthState = { ...DEFAULT_GOOGLE_AUTH_STATE };
let googleAuthBusy = false;
let outlookScanCapabilityState = { ...DEFAULT_OUTLOOK_SCAN_CAPABILITY };
let emailIntakeBusy = false;
let outlookScanState = { ...DEFAULT_OUTLOOK_SCAN_STATE };
let cloudSyncState = { ...DEFAULT_CLOUD_SYNC_STATE };
let cloudSyncConfig = null;
let firebaseLoadPromise = null;
let firebaseAppInstance = null;
let firebaseAuthInstance = null;
let firebaseFirestoreInstance = null;
let cloudSyncInitPromise = null;
let cloudSyncUnsubscribers = [];
let cloudSyncPushTimers = {};
let cloudSyncTimesheetsPushTimer = null;
let cloudSyncApplyDepth = 0;
let cloudSyncRemoteTimesheetMeta = { updatedAt: "", knownWeeks: [] };
let localSyncTimestamps = {
  settings: "",
  tasks: "",
  notes: "",
  templates: "",
  checklists: "",
  timesheets: "",
};
let lastCloudComparableFingerprints = {
  settings: "",
  tasks: "",
  notes: "",
  templates: "",
  checklists: "",
  timesheets: "",
};
let deliverablesFilter = "active";
let separateDeliverableCompletionGroups = true;
let groupDeliverablesByProject = false;
function syncProjectViewPreferencesFromSettings() {
  separateDeliverableCompletionGroups =
    userSettings.separateDeliverableCompletionGroups !== false;
  groupDeliverablesByProject =
    userSettings.groupDeliverablesByProject === true;
}
let activeNoteTab = null;
let activeNotebookType = "note";
let notesSearchQuery = "";
let checklistSearchQuery = "";
let latestAppUpdate = null;
let currentStatsTimespan = "1Y";
let currentStatsAggregation = "month";
let lightingScheduleProjectIndex = null;
let lightingScheduleProjectQuery = "";
let lightingTemplateQuery = "";
let lightingScheduleSyncStatusMessage = "";
let lightingScheduleSyncPollTimer = null;
let lightingScheduleSyncDirty = false;
let lightingScheduleSyncSaving = false;
let title24ProjectIndex = null;
let title24ProjectQuery = "";
let title24ScopeOptionsDataset = null;
let title24ScopeOptionsLoadError = "";
let title24ScopeOptionsLoadingPromise = null;
let openProjectsFilterDropdown = null;
let headerAccountPopoverOpen = false;
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

function stableStringify(value) {
  if (value == null || typeof value !== "object") {
    return JSON.stringify(value);
  }
  if (Array.isArray(value)) {
    return `[${value.map((item) => stableStringify(item)).join(",")}]`;
  }
  const keys = Object.keys(value).sort();
  return `{${keys
    .map((key) => `${JSON.stringify(key)}:${stableStringify(value[key])}`)
    .join(",")}}`;
}

function deepCloneJson(value, fallback = null) {
  if (value == null) return fallback;
  try {
    return JSON.parse(JSON.stringify(value));
  } catch (e) {
    return fallback;
  }
}

function normalizeCloudSyncSettings(raw = {}) {
  return {
    ...DEFAULT_CLOUD_SYNC_SETTINGS,
    ...(raw && typeof raw === "object" ? raw : {}),
    enabled: raw?.enabled === true,
    firebaseUid: String(raw?.firebaseUid || "").trim(),
    lastSyncedAt: String(raw?.lastSyncedAt || "").trim(),
    migrationCompleted: raw?.migrationCompleted === true,
  };
}

function ensureCloudSyncSettingsObject() {
  userSettings.cloudSync = normalizeCloudSyncSettings(userSettings.cloudSync);
  return userSettings.cloudSync;
}

function normalizeIsoTimestamp(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  const parsed = Date.parse(raw);
  return Number.isFinite(parsed) ? new Date(parsed).toISOString() : "";
}

function isIsoAfter(left, right) {
  const leftMs = Date.parse(normalizeIsoTimestamp(left) || "");
  const rightMs = Date.parse(normalizeIsoTimestamp(right) || "");
  if (!Number.isFinite(leftMs)) return false;
  if (!Number.isFinite(rightMs)) return true;
  return leftMs > rightMs;
}

function getLatestIsoTimestamp(values = []) {
  return values.reduce((latest, value) => {
    if (isIsoAfter(value, latest)) return normalizeIsoTimestamp(value);
    return latest;
  }, "");
}

function touchLocalSyncTimestamp(domain, timestamp = new Date().toISOString()) {
  const normalized = normalizeIsoTimestamp(timestamp) || new Date().toISOString();
  localSyncTimestamps[domain] = normalized;
  return normalized;
}

function formatSyncTimestamp(value) {
  const normalized = normalizeIsoTimestamp(value);
  if (!normalized) return "";
  try {
    return new Date(normalized).toLocaleString();
  } catch (e) {
    return normalized;
  }
}

function beginCloudSyncApply() {
  cloudSyncApplyDepth += 1;
}

function endCloudSyncApply() {
  cloudSyncApplyDepth = Math.max(0, cloudSyncApplyDepth - 1);
}

function isCloudSyncApplying() {
  return cloudSyncApplyDepth > 0;
}

function isLikelyLocalPath(value) {
  const raw = String(value || "").trim();
  if (!raw) return false;
  return (
    /^file:\/\//i.test(raw) ||
    /^[a-z]:[\\/]/i.test(raw) ||
    /^\\\\/.test(raw)
  );
}

function isLocalOnlyLink(link) {
  const raw = String(link?.raw || link?.url || link || "").trim();
  return isLikelyLocalPath(raw);
}

function normalizeGoogleAuthState(raw = {}) {
  return {
    ...DEFAULT_GOOGLE_AUTH_STATE,
    ...(raw && typeof raw === "object" ? raw : {}),
    signedIn: !!raw?.signedIn,
    provider: String(raw?.provider || "google").trim() || "google",
    email: String(raw?.email || "").trim(),
    displayName: String(raw?.displayName || "").trim(),
    avatarUrl: String(raw?.avatarUrl || "").trim(),
    signedInAt: String(raw?.signedInAt || "").trim(),
    expiresAt: String(raw?.expiresAt || "").trim(),
    hasRefreshToken: !!raw?.hasRefreshToken,
  };
}

function getGoogleAuthDisplayName(state = googleAuthState) {
  return state.displayName || state.email || "Google";
}

function getGoogleAuthInitials(state = googleAuthState) {
  const displayName = String(state?.displayName || "").trim();
  if (displayName) {
    const initials = displayName
      .split(/\s+/)
      .filter(Boolean)
      .slice(0, 2)
      .map((token) => {
        const match = token.match(/[A-Za-z0-9]/);
        return match ? match[0].toUpperCase() : "";
      })
      .join("");
    if (initials) return initials;
  }

  const emailLocal = String(state?.email || "")
    .split("@")[0]
    .trim();
  const emailFallback = (emailLocal.match(/[A-Za-z0-9]/g) || [])
    .slice(0, 2)
    .join("")
    .toUpperCase();
  return emailFallback || "G";
}

function getGoogleAuthAvatarFallback(state = googleAuthState) {
  return getGoogleAuthInitials(state);
}

function applyGoogleAuthAvatar(
  element,
  state = googleAuthState,
  { forceInitials = false } = {}
) {
  if (!element) return;
  const fallback = getGoogleAuthAvatarFallback(state);
  element.textContent = fallback;
  if (!forceInitials && state.avatarUrl) {
    const safeUrl = state.avatarUrl.replace(/'/g, "%27");
    element.style.backgroundImage = `url('${safeUrl}')`;
    element.style.color = "transparent";
  } else {
    element.style.backgroundImage = "";
    element.style.color = "";
  }
}

function updateHeaderAccountPopoverVisibility() {
  const dropdown = document.getElementById("headerAccountDropdown");
  const popover = document.getElementById("headerAccountPopover");
  const headerBtn = document.getElementById("headerGoogleAuthBtn");
  const isOpen = googleAuthState.signedIn && !googleAuthBusy && headerAccountPopoverOpen;

  headerAccountPopoverOpen = isOpen;
  dropdown?.classList.toggle("open", isOpen);
  if (popover) {
    popover.hidden = !isOpen;
  }
  if (headerBtn) {
    headerBtn.setAttribute("aria-expanded", String(isOpen));
  }
}

function setHeaderAccountPopoverOpen(isOpen, { focusTrigger = false } = {}) {
  headerAccountPopoverOpen = !!isOpen;
  updateHeaderAccountPopoverVisibility();
  if (focusTrigger) {
    document.getElementById("headerGoogleAuthBtn")?.focus();
  }
}

function toggleHeaderAccountPopover() {
  setHeaderAccountPopoverOpen(!headerAccountPopoverOpen);
}

function renderGoogleAuthUi() {
  const headerBtn = document.getElementById("headerGoogleAuthBtn");
  const headerLabel = document.getElementById("headerGoogleAuthLabel");
  const headerMeta = document.getElementById("headerGoogleAuthMeta");
  const headerAvatar = document.getElementById("headerGoogleAuthAvatar");
  const statusEl = document.getElementById("settings_googleAuthStatus");
  const detailsEl = document.getElementById("settings_googleAuthDetails");
  const settingsActionBtn = document.getElementById(
    "settings_googleAuthActionBtn"
  );
  const signOutBtn = document.getElementById("googleSignOutBtn");
  const headerPopoverName = document.getElementById("headerAccountPopoverName");
  const headerPopoverEmail = document.getElementById("headerAccountPopoverEmail");
  const headerPopoverAvatar = document.getElementById("headerAccountPopoverAvatar");
  const headerPopoverSignOutBtn = document.getElementById(
    "headerAccountSignOutBtn"
  );
  const accountName = document.getElementById("googleAccountName");
  const accountEmail = document.getElementById("googleAccountEmail");
  const accountAvatar = document.getElementById("googleAccountAvatar");
  const displayName = getGoogleAuthDisplayName();
  const isSignedIn = googleAuthState.signedIn;
  const primaryLabel = googleAuthBusy
    ? "Connecting..."
    : isSignedIn
      ? displayName
      : "Sign in";
  const secondaryLabel = googleAuthBusy
    ? "Check your browser"
    : isSignedIn
      ? googleAuthState.email && googleAuthState.email !== displayName
        ? googleAuthState.email
        : "Signed in"
      : "Google account";

  if (headerBtn) {
    headerBtn.disabled = googleAuthBusy;
    headerBtn.dataset.signedIn = isSignedIn ? "true" : "false";
    headerBtn.classList.toggle("is-compact", isSignedIn);
    headerBtn.setAttribute(
      "aria-label",
      isSignedIn
        ? `Open Google account menu for ${displayName}`
        : "Sign in with Google"
    );
    if (isSignedIn) {
      headerBtn.setAttribute("aria-haspopup", "dialog");
      headerBtn.setAttribute("aria-controls", "headerAccountPopover");
    } else {
      headerBtn.removeAttribute("aria-haspopup");
      headerBtn.removeAttribute("aria-controls");
    }
    headerBtn.title = googleAuthBusy
      ? "Completing Google sign-in"
      : isSignedIn
        ? `Signed in as ${displayName}`
        : "Sign in with Google";
  }
  if (headerLabel) {
    headerLabel.textContent = primaryLabel;
  }
  if (headerMeta) {
    headerMeta.textContent = secondaryLabel;
  }
  applyGoogleAuthAvatar(headerAvatar, googleAuthState, {
    forceInitials: isSignedIn,
  });
  if (headerPopoverName) {
    headerPopoverName.textContent = isSignedIn ? displayName : "Not signed in";
  }
  if (headerPopoverEmail) {
    headerPopoverEmail.textContent = isSignedIn
      ? googleAuthState.email || "Signed in with Google."
      : "Sign in to connect your Google account.";
  }
  applyGoogleAuthAvatar(headerPopoverAvatar, googleAuthState);
  if (headerPopoverSignOutBtn) {
    headerPopoverSignOutBtn.disabled = googleAuthBusy || !isSignedIn;
    headerPopoverSignOutBtn.textContent = googleAuthBusy
      ? "Signing out..."
      : "Sign out";
  }

  if (statusEl) {
    statusEl.textContent = isSignedIn
      ? `Signed in as ${displayName}`
      : "Not signed in";
  }
  if (detailsEl) {
    detailsEl.textContent = isSignedIn
      ? googleAuthState.email || "Google account connected."
      : "Sign in with Google. Configure GOOGLE_OAUTH_CLIENT_ID and, if your Google OAuth client requires it, GOOGLE_CLIENT_SECRET.";
  }
  if (settingsActionBtn) {
    settingsActionBtn.disabled = googleAuthBusy;
    settingsActionBtn.textContent = googleAuthBusy
      ? "Working..."
      : isSignedIn
        ? "Account"
        : "Sign in";
  }
  if (signOutBtn) {
    signOutBtn.disabled = googleAuthBusy || !isSignedIn;
    signOutBtn.style.display = isSignedIn ? "inline-flex" : "none";
    signOutBtn.textContent = googleAuthBusy ? "Signing out..." : "Sign out";
  }
  if (accountName) {
    accountName.textContent = isSignedIn
      ? displayName
      : "Not signed in";
  }
  if (accountEmail) {
    accountEmail.textContent = isSignedIn
      ? googleAuthState.email || "Signed in with Google."
      : "Sign in to connect your Google account.";
  }
  applyGoogleAuthAvatar(accountAvatar, googleAuthState);
  if (!isSignedIn || googleAuthBusy) {
    headerAccountPopoverOpen = false;
  }
  updateHeaderAccountPopoverVisibility();
  renderCloudSyncUi();
}

async function loadGoogleAuthState({ silent = false } = {}) {
  if (!window.pywebview?.api?.get_google_auth_state) {
    googleAuthState = normalizeGoogleAuthState();
    renderGoogleAuthUi();
    return googleAuthState;
  }
  try {
    const response = await window.pywebview.api.get_google_auth_state();
    if (response?.status !== "success") {
      throw new Error(response?.message || "Failed to load Google sign-in state.");
    }
    googleAuthState = normalizeGoogleAuthState(response.auth);
  } catch (e) {
    console.warn("Failed to load Google auth state:", e);
    googleAuthState = normalizeGoogleAuthState();
    if (!silent) {
      toast("Could not load Google sign-in state.");
    }
  }
  renderGoogleAuthUi();
  return googleAuthState;
}

function openGoogleAccountDialog() {
  const dialog = document.getElementById("googleAccountDlg");
  setHeaderAccountPopoverOpen(false);
  renderGoogleAuthUi();
  showDialog(dialog);
}

function handleHeaderGoogleAuthAction() {
  if (googleAuthBusy) return;
  if (googleAuthState.signedIn) {
    toggleHeaderAccountPopover();
    return;
  }
  handleGoogleSignIn();
}

function handleGoogleAuthAction() {
  if (googleAuthBusy) return;
  if (googleAuthState.signedIn) {
    openGoogleAccountDialog();
    return;
  }
  handleGoogleSignIn();
}

async function handleGoogleSignIn() {
  if (googleAuthBusy || !window.pywebview?.api?.sign_in_with_google) return;
  googleAuthBusy = true;
  renderGoogleAuthUi();
  try {
    const response = await window.pywebview.api.sign_in_with_google();
    if (response?.status === "success") {
      googleAuthState = normalizeGoogleAuthState(response.auth);
      await loadUserSettings();
      const verifiedAuthState = await loadGoogleAuthState({ silent: true });
      if (!verifiedAuthState.signedIn) {
        toast(
          "Google sign-in completed in the browser, but the app could not confirm your signed-in state."
        );
        return;
      }
      googleAuthBusy = false;
      renderGoogleAuthUi();
      showAppLoader();
      try {
        await bootstrapCloudSync({
          session: response.syncSession || null,
          silent: false,
        });
      } finally {
        hideAppLoader();
      }
      if (cloudSyncState.enabled) {
        toast(`Signed in as ${getGoogleAuthDisplayName()}. Cloud sync is active.`);
      } else if (cloudSyncState.error) {
        toast(
          `Signed in as ${getGoogleAuthDisplayName()}, but cloud sync failed: ${cloudSyncState.error}`
        );
      } else if (!cloudSyncState.configured) {
        toast(
          `Signed in as ${getGoogleAuthDisplayName()}. Firebase sync is not configured, so data stays local.`
        );
      } else {
        toast(`Signed in as ${getGoogleAuthDisplayName()}.`);
      }
      return;
    }
    if (response?.status === "cancelled") {
      toast(response.message || "Google sign-in was cancelled.");
      return;
    }
    throw new Error(response?.message || "Google sign-in failed.");
  } catch (e) {
    console.warn("Google sign-in failed:", e);
    toast(e?.message || "Google sign-in failed.");
  } finally {
    googleAuthBusy = false;
    renderGoogleAuthUi();
  }
}

async function handleGoogleSignOut() {
  if (googleAuthBusy || !window.pywebview?.api?.sign_out_google) return;
  setHeaderAccountPopoverOpen(false);
  googleAuthBusy = true;
  renderGoogleAuthUi();
  try {
    const response = await window.pywebview.api.sign_out_google();
    if (response?.status !== "success") {
      throw new Error(response?.message || "Google sign-out failed.");
    }
    closeDlg("googleAccountDlg");
    await signOutCloud({ preserveMetadata: true });
    await loadUserSettings();
    googleAuthState = normalizeGoogleAuthState(response.auth);
    renderGoogleAuthUi();
    toast("Signed out of Google.");
  } catch (e) {
    console.warn("Google sign-out failed:", e);
    toast(e?.message || "Google sign-out failed.");
  } finally {
    googleAuthBusy = false;
    renderGoogleAuthUi();
  }
}

function normalizeEmailIntakeMode(value = "") {
  return String(value || "").trim().toLowerCase() === "scan" ? "scan" : "paste";
}

function normalizeOutlookScanCapability(raw = {}) {
  return {
    ...DEFAULT_OUTLOOK_SCAN_CAPABILITY,
    ...(raw && typeof raw === "object" ? raw : {}),
    loaded: raw?.loaded !== undefined ? !!raw.loaded : true,
    loading: !!raw?.loading,
    desktopAvailable: !!raw?.desktopAvailable,
    desktopReason: String(raw?.desktopReason || "").trim(),
  };
}

function getOutlookScanSourceLabel(source = "") {
  const normalized = String(source || "").trim().toLowerCase();
  if (normalized === "desktop-outlook") return "Desktop Outlook";
  return "Outlook";
}

function getOutlookScanVisibleSuggestions() {
  const dismissed = new Set(outlookScanState.dismissedKeys || []);
  return (outlookScanState.suggestions || []).filter(
    (suggestion) => suggestion && !dismissed.has(suggestion.key)
  );
}

function setOutlookScanProgress(
  payload = {},
  { reset = false, appendLog = true } = {}
) {
  const previous = reset
    ? createEmptyOutlookScanProgress()
    : normalizeOutlookScanProgress(outlookScanState.progress || {});
  const next = normalizeOutlookScanProgress(
    {
      ...(payload || {}),
      receivedAt: payload?.receivedAt || new Date().toISOString(),
    },
    previous
  );
  outlookScanState.progress = next;
  if (reset) {
    outlookScanState.progressLog = [];
  }
  if (appendLog && (next.stage || next.message)) {
    outlookScanState.progressLog = [
      ...(Array.isArray(outlookScanState.progressLog)
        ? outlookScanState.progressLog
        : []),
      { ...next },
    ].slice(-100);
  }
  return next;
}

function formatOutlookScanTimeframeLabel(timeframe = "") {
  return String(timeframe || "").trim().toLowerCase() === "month"
    ? "This month"
    : "This week";
}

function formatOutlookScanSelectedDayLabel(scanDate = "", timeframe = "") {
  const normalized = normalizeOutlookScanDateInput(scanDate || "");
  if (normalized) {
    return parseOutlookScanDateValue(normalized).toLocaleDateString([], {
      year: "numeric",
      month: "short",
      day: "numeric",
    });
  }
  return formatOutlookScanTimeframeLabel(timeframe);
}

function formatOutlookScanStageLabel(stage = "") {
  const key = String(stage || "").trim().toLowerCase();
  const labels = {
    starting: "Starting scan",
    listing: "Loading inbox",
    fallback: "Switching source",
    hydrating: "Reading emails",
    preparing_ai: "Preparing AI review",
    reviewing_ai: "Reviewing with AI",
    matching: "Matching suggestions",
    done: "Scan complete",
    error: "Scan failed",
  };
  return labels[key] || "Outlook scan";
}

function buildOutlookScanProgressParts(progress = {}) {
  const parts = [];
  const totalEmails = normalizeOutlookScanCount(progress?.totalEmails);
  const processedEmails = normalizeOutlookScanCount(progress?.processedEmails);
  const includedEmails = normalizeOutlookScanCount(progress?.includedEmails);
  const skippedEmails = normalizeOutlookScanCount(progress?.skippedEmails);
  const deliverablesInPeriod = normalizeOutlookScanCount(
    progress?.deliverablesInPeriod
  );
  const relevantEmails = normalizeOutlookScanCount(progress?.relevantEmails);
  const threadsDetected = normalizeOutlookScanCount(progress?.threadsDetected);
  const dedupedEmailCount = normalizeOutlookScanCount(progress?.dedupedEmailCount);
  const dedupeSkippedEmailCount = normalizeOutlookScanCount(
    progress?.dedupeSkippedEmailCount
  );

  if (totalEmails !== null && processedEmails !== null && totalEmails > 0) {
    parts.push(
      `${Math.min(processedEmails, totalEmails)} of ${totalEmails} emails processed`
    );
  } else if (totalEmails !== null) {
    parts.push(`${totalEmails} email${totalEmails === 1 ? "" : "s"} found`);
  } else if (processedEmails !== null) {
    parts.push(
      `${processedEmails} email${processedEmails === 1 ? "" : "s"} processed`
    );
  }
  if (includedEmails !== null) {
    parts.push(
      `${includedEmails} email${includedEmails === 1 ? "" : "s"} included in AI review`
    );
  }
  if (skippedEmails !== null && skippedEmails > 0) {
    parts.push(`${skippedEmails} skipped`);
  }
  if (deliverablesInPeriod !== null) {
    parts.push(
      `${deliverablesInPeriod} deliverable${deliverablesInPeriod === 1 ? "" : "s"} on day`
    );
  }
  if (relevantEmails !== null && relevantEmails > 0) {
    parts.push(
      `${relevantEmails} relevant email${relevantEmails === 1 ? "" : "s"} found`
    );
  }
  if (threadsDetected !== null && threadsDetected > 0) {
    parts.push(
      `${threadsDetected} thread${threadsDetected === 1 ? "" : "s"} detected`
    );
  }
  if (dedupedEmailCount !== null && dedupedEmailCount > 0) {
    parts.push(
      `${dedupedEmailCount} email${dedupedEmailCount === 1 ? "" : "s"} shortened`
    );
  }
  if (dedupeSkippedEmailCount !== null && dedupeSkippedEmailCount > 0) {
    parts.push(
      `${dedupeSkippedEmailCount} duplicate email${dedupeSkippedEmailCount === 1 ? "" : "s"} skipped`
    );
  }
  return parts;
}

function getOutlookScanProgressSummary(progress = {}) {
  return (
    String(progress?.message || "").trim() ||
    formatOutlookScanStageLabel(progress?.stage)
  );
}

function formatOutlookScanLogTime(value = "") {
  const date = value ? new Date(value) : null;
  if (!(date instanceof Date) || isNaN(date)) return "";
  return date.toLocaleTimeString([], {
    hour: "numeric",
    minute: "2-digit",
    second: "2-digit",
  });
}

function buildOutlookScanReportModel() {
  const progress = normalizeOutlookScanProgress(outlookScanState.progress || {});
  const log = Array.isArray(outlookScanState.progressLog)
    ? outlookScanState.progressLog.map((entry) =>
        normalizeOutlookScanProgress(
          entry || {},
          createEmptyOutlookScanProgress()
        )
      )
    : [];
  const stats =
    outlookScanState.reportStats && typeof outlookScanState.reportStats === "object"
      ? outlookScanState.reportStats
      : {};
  const lastResult =
    outlookScanState.lastResult && typeof outlookScanState.lastResult === "object"
      ? outlookScanState.lastResult
      : null;
  return {
    source:
      String(stats.source || lastResult?.source || progress.source || "").trim(),
    scanDate: normalizeOutlookScanDateInput(
      stats.scanDate ||
        lastResult?.scanDate ||
        progress.scanDate ||
        outlookScanState.scanDate ||
        ""
    ),
    timeframe: String(
      stats.timeframe ||
        lastResult?.timeframe ||
        progress.timeframe ||
        (outlookScanState.scanDate ? "" : "week") ||
        "week"
    )
      .trim()
      .toLowerCase(),
    fallbackUsed: log.some((entry) => entry.stage === "fallback"),
    totalEmails:
      normalizeOutlookScanCount(progress.totalEmails) ??
      normalizeOutlookScanCount(lastResult?.scannedCount),
    processedEmails:
      normalizeOutlookScanCount(progress.processedEmails) ??
      normalizeOutlookScanCount(lastResult?.scannedCount),
    includedEmails:
      normalizeOutlookScanCount(progress.includedEmails) ??
      normalizeOutlookScanCount(lastResult?.emailsIncludedCount),
    skippedEmails:
      normalizeOutlookScanCount(progress.skippedEmails) ??
      normalizeOutlookScanCount(stats.skippedCount) ??
      normalizeOutlookScanCount(
        Array.isArray(lastResult?.skippedMessages)
          ? lastResult.skippedMessages.length
          : null
      ),
    deliverablesInPeriod:
      normalizeOutlookScanCount(progress.deliverablesInPeriod) ??
      normalizeOutlookScanCount(stats.deliverablesInPeriod) ??
      normalizeOutlookScanCount(lastResult?.deliverablesIncludedCount),
    relevantEmails:
      normalizeOutlookScanCount(progress.relevantEmails) ??
      normalizeOutlookScanCount(lastResult?.relevantEmailCount),
    threadsDetected:
      normalizeOutlookScanCount(progress.threadsDetected) ??
      normalizeOutlookScanCount(stats.threadsDetected) ??
      normalizeOutlookScanCount(lastResult?.threadsDetected),
    dedupedEmailCount:
      normalizeOutlookScanCount(progress.dedupedEmailCount) ??
      normalizeOutlookScanCount(stats.dedupedEmailCount) ??
      normalizeOutlookScanCount(lastResult?.dedupedEmailCount),
    dedupeSkippedEmailCount:
      normalizeOutlookScanCount(progress.dedupeSkippedEmailCount) ??
      normalizeOutlookScanCount(stats.dedupeSkippedEmailCount) ??
      normalizeOutlookScanCount(lastResult?.dedupeSkippedEmailCount),
    promptTruncated: !!(lastResult?.promptTruncated || stats.promptTruncated),
    suggestionCount:
      normalizeOutlookScanCount(stats.suggestionCount) ??
      normalizeOutlookScanCount(
        Array.isArray(lastResult?.suggestions) ? lastResult.suggestions.length : 0
      ) ??
      0,
    errorMessage: String(
      stats.errorMessage || lastResult?.message || ""
    ).trim(),
    log,
  };
}

window.updateOutlookScanProgress = function (payload = {}) {
  setOutlookScanProgress(payload, { appendLog: true });
  const dialog = document.getElementById("outlookScanDlg");
  if (dialog?.open) {
    renderOutlookScanUi();
  }
};

function renderOutlookScanUi() {
  const toolbarBtn = document.getElementById("outlookScanBtn");
  const pasteModeInput = document.getElementById("emailIntakeModePaste");
  const scanModeInput = document.getElementById("emailIntakeModeScan");
  const pastePanel = document.getElementById("emailIntakePastePanel");
  const scanPanel = document.getElementById("emailIntakeScanPanel");
  const capabilityStatusEl = document.getElementById("outlookScanCapabilityStatus");
  const capabilityDetailsEl = document.getElementById(
    "outlookScanCapabilityDetails"
  );
  const emailArea = document.getElementById("emailArea");
  const aiSpinner = document.getElementById("aiSpinner");
  const processBtn = document.getElementById("btnProcessEmail");
  const scanDateInput = document.getElementById("outlookScanDate");
  const progressEl = document.getElementById("outlookScanProgress");
  const runBtn = document.getElementById("outlookScanRunBtn");
  const summaryEl = document.getElementById("outlookScanSummary");
  const metaEl = document.getElementById("outlookScanMeta");
  const reportEl = document.getElementById("outlookScanReport");
  const reportSummaryEl = document.getElementById("outlookScanReportSummary");
  const reportLogEl = document.getElementById("outlookScanReportLog");
  const suggestionsEl = document.getElementById("outlookScanSuggestions");
  const skippedEl = document.getElementById("outlookScanSkipped");
  const emptyEl = document.getElementById("outlookScanEmpty");

  const mode = normalizeEmailIntakeMode(outlookScanState.mode);
  const capability = normalizeOutlookScanCapability(outlookScanCapabilityState);
  const desktopAvailable = !!capability.desktopAvailable;
  const desktopReason = String(capability.desktopReason || "").trim();
  const scanBusy = outlookScanState.busy;
  const busy = scanBusy || capability.loading;
  const visibleSuggestions = getOutlookScanVisibleSuggestions();
  const skipped = Array.isArray(outlookScanState.skipped)
    ? outlookScanState.skipped
    : [];
  const lastResult = outlookScanState.lastResult;
  const progress = normalizeOutlookScanProgress(outlookScanState.progress || {});
  const reportModel = buildOutlookScanReportModel();
  const selectedDayLabel = formatOutlookScanSelectedDayLabel(
    reportModel.scanDate,
    reportModel.timeframe
  );
  const showReport =
    !scanBusy &&
    !!(reportModel.log.length || reportModel.errorMessage || lastResult);

  if (toolbarBtn) {
    toolbarBtn.disabled = false;
    toolbarBtn.title = scanBusy
      ? "Outlook day scan is running"
      : emailIntakeBusy
        ? "AI email intake is running"
        : capability.loading
          ? "Checking Desktop Outlook availability"
          : "Email intake";
  }
  if (pasteModeInput) {
    pasteModeInput.checked = mode === "paste";
    pasteModeInput.disabled = emailIntakeBusy || scanBusy;
  }
  if (scanModeInput) {
    scanModeInput.checked = mode === "scan";
    scanModeInput.disabled = emailIntakeBusy || scanBusy;
  }
  if (pastePanel) {
    pastePanel.hidden = mode !== "paste";
  }
  if (scanPanel) {
    scanPanel.hidden = mode !== "scan";
  }
  if (capabilityStatusEl) {
    capabilityStatusEl.textContent = capability.loading
      ? "Checking Desktop Outlook..."
      : desktopAvailable
        ? "Desktop Outlook ready"
        : "Desktop Outlook unavailable";
  }
  if (capabilityDetailsEl) {
    capabilityDetailsEl.textContent = capability.loading
      ? "Checking whether Desktop Outlook is available on this machine."
      : desktopAvailable
        ? "Scan a selected day using the installed Outlook desktop app on this machine."
        : desktopReason || "Desktop Outlook is unavailable on this machine.";
  }
  if (emailArea) {
    emailArea.disabled = emailIntakeBusy;
  }
  if (aiSpinner) {
    aiSpinner.hidden = !emailIntakeBusy;
  }
  if (processBtn) {
    processBtn.disabled = emailIntakeBusy;
    processBtn.textContent = emailIntakeBusy
      ? "Processing..."
      : "Process with AI";
  }
  if (scanDateInput) {
    scanDateInput.value = normalizeOutlookScanDateInput(outlookScanState.scanDate);
    scanDateInput.max = getTodayLocalDateInputValue();
    scanDateInput.disabled = busy;
  }
  if (runBtn) {
    runBtn.disabled = busy || !desktopAvailable;
    runBtn.textContent = scanBusy
      ? "Scanning..."
      : capability.loading
        ? "Checking..."
        : "Scan day";
  }
  if (progressEl) {
    progressEl.classList.toggle("is-active", scanBusy || progress.active);
    progressEl.classList.toggle(
      "is-error",
      !scanBusy && String(lastResult?.status || "").trim().toLowerCase() === "error"
    );
  }
  if (summaryEl) {
    if (scanBusy || progress.active) {
      summaryEl.textContent = getOutlookScanProgressSummary(progress);
    } else if (
      String(lastResult?.status || "").trim().toLowerCase() === "error"
    ) {
      summaryEl.textContent = reportModel.errorMessage || "Outlook scan failed.";
    } else if (!lastResult) {
      summaryEl.textContent = "No Outlook day scan has been run yet.";
    } else {
      summaryEl.textContent =
        `${visibleSuggestions.length} suggestion` +
        `${visibleSuggestions.length === 1 ? "" : "s"}, ` +
        `${skipped.length} skipped item${skipped.length === 1 ? "" : "s"}`;
    }
  }
  if (metaEl) {
    if (scanBusy || progress.active) {
      const progressParts = buildOutlookScanProgressParts(progress);
      metaEl.textContent =
        progressParts.join(" · ") ||
        "This can take a bit if there are many emails on the selected day.";
    } else if (
      String(lastResult?.status || "").trim().toLowerCase() === "error"
    ) {
      const parts = [];
      if (reportModel.source) {
        parts.push(`Source: ${getOutlookScanSourceLabel(reportModel.source)}`);
      }
      parts.push(`Day: ${selectedDayLabel}`);
      if (reportModel.promptTruncated) {
        parts.push("Prompt was trimmed to keep the scan bounded.");
      }
      if (reportModel.errorMessage) {
        parts.push(reportModel.errorMessage);
      }
      metaEl.textContent = parts.join(" · ");
    } else if (!lastResult) {
      metaEl.textContent = desktopAvailable
        ? "Choose a day, then run a scan."
        : capability.loading
          ? "Checking whether Desktop Outlook is available."
          : desktopReason || "Desktop Outlook is unavailable on this machine.";
    } else {
      const deliverableCount =
        reportModel.deliverablesInPeriod ?? lastResult.deliverablesIncludedCount ?? 0;
      const parts = [
        `Scanned ${Number(lastResult.scannedCount || 0)} email${Number(lastResult.scannedCount || 0) === 1 ? "" : "s"}`,
        `Included ${Number(lastResult.emailsIncludedCount || 0)} email${Number(lastResult.emailsIncludedCount || 0) === 1 ? "" : "s"} and ${Number(deliverableCount || 0)} current deliverable${Number(deliverableCount || 0) === 1 ? "" : "s"} in one AI review`,
        `Day: ${selectedDayLabel}`,
      ];
      if (lastResult?.source) {
        parts.push(`Source: ${getOutlookScanSourceLabel(lastResult.source)}`);
      }
      if (lastResult.promptTruncated) {
        parts.push("Prompt was trimmed to keep the scan bounded.");
      } else if (lastResult.truncated) {
        parts.push("Results were truncated to keep the scan bounded.");
      }
      metaEl.textContent = parts.join(" · ");
    }
  }
  if (reportEl) {
    reportEl.hidden = !showReport;
    if (!showReport) {
      reportEl.open = false;
    } else if (reportEl.open !== !!outlookScanState.reportOpen) {
      reportEl.open = !!outlookScanState.reportOpen;
    }
  }
  if (reportSummaryEl) {
    reportSummaryEl.innerHTML = "";
    if (showReport) {
      const rows = [];
      if (reportModel.source) {
        rows.push(["Source", getOutlookScanSourceLabel(reportModel.source)]);
      }
      rows.push(["Day", selectedDayLabel]);
      if (reportModel.totalEmails !== null) {
        rows.push(["Emails found", `${reportModel.totalEmails}`]);
      }
      if (reportModel.processedEmails !== null) {
        rows.push(["Emails processed", `${reportModel.processedEmails}`]);
      }
      if (reportModel.includedEmails !== null) {
        rows.push(["Emails in AI review", `${reportModel.includedEmails}`]);
      }
      if (reportModel.skippedEmails !== null) {
        rows.push(["Skipped emails", `${reportModel.skippedEmails}`]);
      }
      if (reportModel.deliverablesInPeriod !== null) {
        rows.push([
          "Deliverables on day",
          `${reportModel.deliverablesInPeriod}`,
        ]);
      }
      if (reportModel.relevantEmails !== null) {
        rows.push(["Relevant emails", `${reportModel.relevantEmails}`]);
      }
      if (reportModel.threadsDetected !== null) {
        rows.push(["Threads detected", `${reportModel.threadsDetected}`]);
      }
      if (reportModel.dedupedEmailCount !== null) {
        rows.push(["Emails shortened", `${reportModel.dedupedEmailCount}`]);
      }
      if (reportModel.dedupeSkippedEmailCount !== null) {
        rows.push([
          "Duplicate emails skipped",
          `${reportModel.dedupeSkippedEmailCount}`,
        ]);
      }
      rows.push(["Prompt trimmed", reportModel.promptTruncated ? "Yes" : "No"]);
      rows.push(["Suggestions", `${reportModel.suggestionCount}`]);
      if (reportModel.errorMessage) {
        rows.push(["Error", reportModel.errorMessage]);
      }
      rows.forEach(([label, value]) => {
        reportSummaryEl.appendChild(
          el("div", { className: "outlook-scan-report-row" }, [
            el("div", {
              className: "outlook-scan-report-label tiny muted",
              textContent: label,
            }),
            el("div", {
              className: "outlook-scan-report-value",
              textContent: value,
            }),
          ])
        );
      });
    }
  }
  if (reportLogEl) {
    reportLogEl.innerHTML = "";
    if (showReport) {
      if (reportModel.log.length) {
        reportModel.log.forEach((entry) => {
          const entryParts = buildOutlookScanProgressParts(entry);
          reportLogEl.appendChild(
            el("div", { className: "outlook-scan-log-item" }, [
              el("div", { className: "outlook-scan-log-head" }, [
                el("div", {
                  className: "outlook-scan-log-stage",
                  textContent: formatOutlookScanStageLabel(entry.stage),
                }),
                el("div", {
                  className: "tiny muted",
                  textContent: formatOutlookScanLogTime(entry.receivedAt),
                }),
              ]),
              el("div", {
                className: "outlook-scan-log-message",
                textContent:
                  String(entry.message || "").trim() ||
                  formatOutlookScanStageLabel(entry.stage),
              }),
              entryParts.length
                ? el("div", {
                    className: "tiny muted",
                    textContent: entryParts.join(" · "),
                  })
                : null,
            ].filter(Boolean))
          );
        });
      } else {
        reportLogEl.appendChild(
          el("div", {
            className: "tiny muted",
            textContent: "No scan events were recorded for this run.",
          })
        );
      }
    }
  }

  if (suggestionsEl) {
    suggestionsEl.innerHTML = "";
    visibleSuggestions.forEach((suggestion) => {
      const projectLabel =
        suggestion.projectName ||
        suggestion.projectId ||
        `Project ${suggestion.projectIndex + 1}`;
      const metaParts = [];
      if (suggestion.due) {
        metaParts.push(`Due ${humanDate(suggestion.due) || suggestion.due}`);
      }
      metaParts.push(
        `${suggestion.relatedMessages.length} related email${suggestion.relatedMessages.length === 1 ? "" : "s"}`
      );
      const card = el("div", { className: "outlook-scan-card" }, [
        el("div", { className: "outlook-scan-card-head" }, [
          el("div", {
            className: "outlook-scan-card-project",
            textContent: projectLabel,
          }),
          el("div", {
            className: "outlook-scan-card-deliverable",
            textContent: suggestion.deliverableName || "Deliverable",
          }),
        ]),
        el("div", {
          className: "outlook-scan-card-meta tiny muted",
          textContent: metaParts.join(" · "),
        }),
        el("div", {
          className: "outlook-scan-card-notes",
          textContent: suggestion.notes || "No additional notes.",
        }),
        el(
          "div",
          { className: "outlook-scan-related-list" },
          suggestion.relatedMessages.slice(0, 3).map((message) =>
            el("button", {
              className: "btn ghost tiny",
              type: "button",
              textContent: message.subject || "Open email",
              onclick: () => openOutlookScanMessage(message),
            })
          )
        ),
        el("div", { className: "outlook-scan-card-actions" }, [
          el("button", {
            className: "btn tiny",
            type: "button",
            textContent: "Dismiss",
            onclick: () => dismissOutlookScanSuggestion(suggestion.key),
          }),
          el("button", {
            className: "btn-primary tiny",
            type: "button",
            textContent: "Add deliverable",
            onclick: () => acceptOutlookScanSuggestion(suggestion.key),
          }),
        ]),
      ]);
      suggestionsEl.appendChild(card);
    });
  }

  if (skippedEl) {
    skippedEl.innerHTML = "";
    skipped.slice(0, 20).forEach((item) => {
      const subject = item?.message?.subject || "Untitled email";
      skippedEl.appendChild(
        el("div", { className: "outlook-scan-skipped-item" }, [
          el("div", {
            className: "outlook-scan-skipped-subject",
            textContent: subject,
          }),
          el("div", {
            className: "tiny muted",
            textContent: item.reason || "Skipped",
          }),
        ])
      );
    });
  }

  if (emptyEl) {
    emptyEl.hidden =
      scanBusy || showReport || visibleSuggestions.length > 0 || skipped.length > 0;
  }
}

async function loadOutlookScanCapability({ silent = false } = {}) {
  if (!window.pywebview?.api?.get_outlook_scan_capability) {
    outlookScanCapabilityState = normalizeOutlookScanCapability({
      loaded: true,
      desktopAvailable: false,
      desktopReason: "Outlook day scan is unavailable in this environment.",
    });
    renderOutlookScanUi();
    return outlookScanCapabilityState;
  }
  outlookScanCapabilityState = normalizeOutlookScanCapability({
    ...outlookScanCapabilityState,
    loaded: true,
    loading: true,
  });
  renderOutlookScanUi();
  try {
    const response = await window.pywebview.api.get_outlook_scan_capability();
    if (response?.status !== "success") {
      throw new Error(
        response?.message || "Failed to load Desktop Outlook scan capability."
      );
    }
    outlookScanCapabilityState = normalizeOutlookScanCapability({
      loaded: true,
      loading: false,
      desktopAvailable: response?.desktopAvailable,
      desktopReason: response?.desktopReason,
    });
  } catch (e) {
    console.warn("Failed to load Outlook scan capability:", e);
    outlookScanCapabilityState = normalizeOutlookScanCapability({
      loaded: true,
      loading: false,
      desktopAvailable: false,
      desktopReason: e?.message || "Desktop Outlook capability could not be loaded.",
    });
    if (!silent) {
      toast("Could not load Desktop Outlook availability.");
    }
  }
  renderOutlookScanUi();
  return outlookScanCapabilityState;
}

function resetOutlookScanState() {
  outlookScanState = {
    ...DEFAULT_OUTLOOK_SCAN_STATE,
    mode: normalizeEmailIntakeMode(outlookScanState.mode),
    scanDate:
      normalizeOutlookScanDateInput(outlookScanState.scanDate) ||
      DEFAULT_OUTLOOK_SCAN_STATE.scanDate,
  };
}

function setEmailIntakeMode(mode = "paste") {
  outlookScanState.mode = normalizeEmailIntakeMode(mode);
  renderOutlookScanUi();
}

async function openOutlookScanDialog(mode = outlookScanState.mode) {
  setEmailIntakeMode(mode);
  const dialog = document.getElementById("outlookScanDlg");
  if (dialog) {
    renderOutlookScanUi();
    showDialog(dialog);
  }
  await loadOutlookScanCapability({ silent: true });
}

async function processEmailIntakePaste() {
  if (emailIntakeBusy || !window.pywebview?.api?.process_email_with_ai) {
    return;
  }
  if (!String(userSettings.apiKey || "").trim()) {
    toast("Setup API Key in Settings first.");
    return;
  }
  const txt = val("emailArea");
  if (!txt) return;
  const AI_EMAIL_TIMEOUT_MS = 120000;
  let timeoutId = null;
  emailIntakeBusy = true;
  renderOutlookScanUi();
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
      const emailField = document.getElementById("emailArea");
      if (emailField) emailField.value = "";
      closeDlg("outlookScanDlg");
      handleAiProjectResult(res.data || {});
      return;
    }
    throw new Error(res?.message || "Failed to process email.");
  } catch (e) {
    toast("AI Error: " + (e?.message || "Unknown error."));
  } finally {
    if (timeoutId) clearTimeout(timeoutId);
    emailIntakeBusy = false;
    renderOutlookScanUi();
  }
}

function setOutlookScanDate(value = "") {
  outlookScanState.scanDate = normalizeOutlookScanDateInput(value);
  renderOutlookScanUi();
}

function getOutlookMessageKey(message = {}) {
  const source = String(message.source || "").trim().toLowerCase();
  const messageId = String(message.id || "").trim().toLowerCase();
  if (messageId) return `${source || "message"}:msg:${messageId}`;
  const internetMessageId = String(message.internetMessageId || "")
    .trim()
    .toLowerCase();
  if (internetMessageId) return `${source || "message"}:internet:${internetMessageId}`;
  return `${source || "message"}:${String(message.webLink || message.url || "")
    .trim()
    .toLowerCase()}`;
}

function getOutlookScanDayRange(scanDate = "") {
  const start = parseOutlookScanDateValue(scanDate);
  start.setHours(0, 0, 0, 0);
  const end = new Date(start);
  end.setHours(23, 59, 59, 999);
  return { start, end };
}

function buildOutlookScanProjectContext(scanDate = "") {
  const { start, end } = getOutlookScanDayRange(scanDate);
  return db
    .map((rawProject) => normalizeProject(rawProject))
    .filter(Boolean)
    .map((project) => {
      const deliverables = getProjectDeliverables(project)
        .filter((deliverable) => {
          const due = parseDueStr(deliverable?.due);
          return due && due >= start && due <= end;
        })
        .map((deliverable) => ({
          name: String(deliverable?.name || "").trim(),
          due: String(deliverable?.due || "").trim(),
          status: String(deliverable?.status || "").trim(),
        }))
        .sort((left, right) => {
          const leftDue = parseDueStr(left?.due);
          const rightDue = parseDueStr(right?.due);
          if (leftDue && rightDue && leftDue.getTime() !== rightDue.getTime()) {
            return leftDue - rightDue;
          }
          if (leftDue && !rightDue) return -1;
          if (!leftDue && rightDue) return 1;
          return `${left.name}|${left.status}`.localeCompare(
            `${right.name}|${right.status}`
          );
        });
      if (!deliverables.length) return null;
      return {
        id: String(project.id || "").trim(),
        name: String(project.name || "").trim(),
        nick: String(project.nick || "").trim(),
        path: String(project.path || "").trim(),
        deliverables,
      };
    })
    .filter(Boolean)
    .sort((left, right) => {
      const leftDue = parseDueStr(left.deliverables[0]?.due);
      const rightDue = parseDueStr(right.deliverables[0]?.due);
      if (leftDue && rightDue && leftDue.getTime() !== rightDue.getTime()) {
        return leftDue - rightDue;
      }
      if (leftDue && !rightDue) return -1;
      if (!leftDue && rightDue) return 1;
      return `${left.id}|${left.name}|${left.path}`.localeCompare(
        `${right.id}|${right.name}|${right.path}`
      );
    });
}

function countOutlookScanProjectContextDeliverables(projectContext = []) {
  return (Array.isArray(projectContext) ? projectContext : []).reduce(
    (total, project) =>
      total +
      (Array.isArray(project?.deliverables) ? project.deliverables.length : 0),
    0
  );
}

function buildOutlookEmailRefFromMessage(message = {}) {
  const source = String(message.source || "").trim().toLowerCase();
  const messageId = String(message.id || "").trim();
  const internetMessageId = String(message.internetMessageId || "").trim();
  if (source === "desktop-outlook" && messageId) {
    return normalizeEmailRef({
      raw: messageId,
      url: "",
      label: String(message.subject || "Email").trim() || "Email",
      source: "outlook-desktop",
      savedAt:
        normalizeIsoTimestamp(message.receivedDateTime) || new Date().toISOString(),
      messageId,
      internetMessageId,
    });
  }
  const webLink = String(message.webLink || message.url || "").trim();
  if (!webLink) return null;
  return normalizeEmailRef({
    raw: webLink,
    url: webLink,
    label: String(message.subject || "Email").trim() || "Email",
    source: "outlook-url",
    savedAt:
      normalizeIsoTimestamp(message.receivedDateTime) || new Date().toISOString(),
    messageId,
    internetMessageId,
  });
}

async function openOutlookScanMessage(message = {}) {
  const ref = buildOutlookEmailRefFromMessage(message);
  if (ref) {
    return openDeliverableEmailRef(ref);
  }
  const webLink = String(message.webLink || message.url || "").trim();
  if (webLink) {
    openExternalUrl(webLink);
    return true;
  }
  toast("Could not open this Outlook message.");
  return false;
}

function normalizeDeliverableComparisonKey(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  return normalizeProjectMatchValue(extractDeliverableName(raw) || raw);
}

function pickEarlierOutlookSuggestionDue(currentDue = "", nextDue = "") {
  const currentDate = parseDueStr(currentDue);
  const nextDate = parseDueStr(nextDue);
  if (!currentDate) return nextDue || currentDue || "";
  if (!nextDate) return currentDue || nextDue || "";
  return nextDate < currentDate ? nextDue : currentDue;
}

function normalizeOutlookSuggestionTasks(tasks = []) {
  const seen = new Set();
  const out = [];
  (Array.isArray(tasks) ? tasks : []).forEach((task) => {
    const normalizedTask = normalizeTask(task);
    const text = String(normalizedTask?.text || task || "").trim();
    const key = text.toLowerCase();
    if (!text || seen.has(key)) return;
    seen.add(key);
    out.push({
      text,
      done: !!normalizedTask?.done,
      links: Array.isArray(normalizedTask?.links) ? normalizedTask.links : [],
    });
  });
  return out;
}

function mergeOutlookSuggestionNotes(currentNotes = "", nextNotes = "") {
  const current = String(currentNotes || "").trim();
  const next = String(nextNotes || "").trim();
  if (!current) return next;
  if (!next) return current;
  if (current.toLowerCase() === next.toLowerCase()) return current;
  return `${current} ${next}`.trim();
}

function mergeOutlookSuggestionTasks(currentTasks = [], nextTasks = []) {
  return normalizeOutlookSuggestionTasks([...(currentTasks || []), ...(nextTasks || [])]);
}

function normalizeOutlookRelatedMessages(messages = []) {
  const seen = new Set();
  return (Array.isArray(messages) ? messages : [])
    .filter((message) => {
      const key = getOutlookMessageKey(message);
      if (!key || seen.has(key)) return false;
      seen.add(key);
      return true;
    })
    .sort(
      (left, right) =>
        Date.parse(normalizeIsoTimestamp(right?.receivedDateTime) || "") -
        Date.parse(normalizeIsoTimestamp(left?.receivedDateTime) || "")
    );
}

function buildOutlookScanSkippedItem(item = {}, fallbackKey = "") {
  const message = item?.message || {};
  return {
    key: String(item?.key || fallbackKey || getOutlookMessageKey(message) || createId("outlook-skip")).trim(),
    message,
    reason: String(item?.reason || "Skipped.").trim() || "Skipped.",
  };
}

function buildOutlookScanSuggestionProject(entry = {}) {
  const project = entry?.project || {};
  return {
    id: String(project.id || entry?.projectId || "").trim(),
    name: String(project.name || entry?.projectName || "").trim(),
    nick: String(project.nick || entry?.projectNick || "").trim(),
    path: String(project.path || entry?.projectPath || "").trim(),
  };
}

function buildOutlookScanSuggestionEmailRefs(messages = []) {
  return normalizeEmailRefs(
    normalizeOutlookRelatedMessages(messages).map((message) =>
      buildOutlookEmailRefFromMessage(message)
    )
  ).slice(0, MAX_DELIVERABLE_EMAIL_REFS);
}

function buildOutlookScanDerivedState(result = null) {
  const rawSuggestions = Array.isArray(result?.suggestions) ? result.suggestions : [];
  const skipped = Array.isArray(result?.skippedMessages)
    ? result.skippedMessages.map((item, index) =>
        buildOutlookScanSkippedItem(item, `backend-skip:${index}`)
      )
    : [];
  const suggestionsByKey = new Map();

  rawSuggestions.forEach((entry, index) => {
    const relatedMessages = normalizeOutlookRelatedMessages(entry?.relatedMessages || []);
    const primaryMessage = relatedMessages[0] || {};
    const projectInfo = buildOutlookScanSuggestionProject(entry);
    const deliverablePayload =
      entry?.deliverable && typeof entry.deliverable === "object" ? entry.deliverable : {};
    const deliverableNameRaw = String(
      deliverablePayload.name || entry?.deliverableName || entry?.deliverable || entry?.name || ""
    ).trim();
    const deliverableName =
      extractDeliverableName(deliverableNameRaw) || deliverableNameRaw;
    const deliverableKey = normalizeDeliverableComparisonKey(deliverableName);
    if (!deliverableKey) {
      skipped.push(
        buildOutlookScanSkippedItem(
          {
            message: primaryMessage,
            reason: "No deliverable was suggested for this email batch.",
          },
          `suggestion:${index}:missing-deliverable`
        )
      );
      return;
    }

    const match = findBestProjectMatch(projectInfo);
    if (!match || !db[match.index]) {
      skipped.push(
        buildOutlookScanSkippedItem(
          {
            message: primaryMessage,
            reason: "No matching existing project was found.",
          },
          `suggestion:${index}:unmatched-project`
        )
      );
      return;
    }

    const project = normalizeProject(db[match.index]);
    const existingDeliverables = new Set(
      getProjectDeliverables(project).map((deliverable) =>
        normalizeDeliverableComparisonKey(deliverable?.name)
      )
    );
    if (existingDeliverables.has(deliverableKey)) {
      skipped.push(
        buildOutlookScanSkippedItem(
          {
            message: primaryMessage,
            reason: "That deliverable already exists on the matched project.",
          },
          `suggestion:${index}:existing-deliverable`
        )
      );
      return;
    }

    const suggestionKey = `project:${match.index}|deliverable:${deliverableKey}`;
    const due = String(deliverablePayload.due || entry?.due || "").trim();
    const tasks = normalizeOutlookSuggestionTasks(
      deliverablePayload.tasks || entry?.tasks || []
    );
    const notes = String(deliverablePayload.notes || entry?.notes || "").trim();
    if (!suggestionsByKey.has(suggestionKey)) {
      suggestionsByKey.set(suggestionKey, {
        key: suggestionKey,
        projectIndex: match.index,
        projectId: String(project.id || "").trim(),
        projectName: String(project.name || "").trim(),
        deliverableName,
        deliverableKey,
        due,
        tasks,
        notes,
        relatedMessages,
      });
      return;
    }

    const existing = suggestionsByKey.get(suggestionKey);
    existing.due = pickEarlierOutlookSuggestionDue(existing.due, due);
    existing.tasks = mergeOutlookSuggestionTasks(existing.tasks, tasks);
    existing.notes = mergeOutlookSuggestionNotes(existing.notes, notes);
    existing.relatedMessages = normalizeOutlookRelatedMessages([
      ...(existing.relatedMessages || []),
      ...relatedMessages,
    ]);
  });

  const suggestions = Array.from(suggestionsByKey.values())
    .map((suggestion) => {
      const relatedMessages = normalizeOutlookRelatedMessages(
        suggestion.relatedMessages || []
      );
      return {
        ...suggestion,
        relatedMessages,
        emailRefs: buildOutlookScanSuggestionEmailRefs(relatedMessages),
      };
    })
    .sort((left, right) =>
      `${left.projectName}|${left.deliverableName}`.localeCompare(
        `${right.projectName}|${right.deliverableName}`
      )
    );

  return { suggestions, skipped };
}

async function runOutlookInboxScan() {
  if (
    outlookScanState.busy ||
    emailIntakeBusy ||
    !window.pywebview?.api?.scan_outlook_inbox
  ) {
    return;
  }
  if (!String(userSettings.apiKey || "").trim()) {
    toast("Setup API Key in Settings first.");
    return;
  }
  const capability = outlookScanCapabilityState.loaded
    ? normalizeOutlookScanCapability(outlookScanCapabilityState)
    : await loadOutlookScanCapability({ silent: true });
  if (capability.loading) {
    return;
  }
  if (!capability.desktopAvailable) {
    toast(
      capability.desktopReason || "Desktop Outlook is unavailable on this machine."
    );
    renderOutlookScanUi();
    return;
  }
  const scanDate = normalizeOutlookScanDateInput(outlookScanState.scanDate);
  outlookScanState.scanDate = scanDate;
  const projectContext = buildOutlookScanProjectContext(scanDate);
  const deliverablesInPeriod =
    countOutlookScanProjectContextDeliverables(projectContext);
  outlookScanState.busy = true;
  outlookScanState.dismissedKeys = [];
  outlookScanState.lastResult = null;
  outlookScanState.suggestions = [];
  outlookScanState.skipped = [];
  outlookScanState.reportOpen = false;
  outlookScanState.reportStats = null;
  setOutlookScanProgress(
    {
      active: true,
      stage: "starting",
      message: "Starting Outlook inbox scan...",
      scanDate,
      deliverablesInPeriod,
    },
    { reset: true, appendLog: false }
  );
  renderOutlookScanUi();
  try {
    const response = await window.pywebview.api.scan_outlook_inbox(
      {
        scanDate,
        projectContext,
      },
      userSettings.apiKey,
      userSettings.userName,
      userSettings.discipline
    );
    const derived = buildOutlookScanDerivedState(response);
    outlookScanState.lastResult = response;
    outlookScanState.suggestions = derived.suggestions;
    outlookScanState.skipped = derived.skipped;
    outlookScanState.reportStats = {
      suggestionCount: derived.suggestions.length,
      skippedCount: derived.skipped.length,
      deliverablesInPeriod:
        normalizeOutlookScanCount(outlookScanState.progress?.deliverablesInPeriod) ??
        deliverablesInPeriod,
      threadsDetected:
        normalizeOutlookScanCount(response?.threadsDetected) ??
        normalizeOutlookScanCount(outlookScanState.progress?.threadsDetected),
      dedupedEmailCount:
        normalizeOutlookScanCount(response?.dedupedEmailCount) ??
        normalizeOutlookScanCount(outlookScanState.progress?.dedupedEmailCount),
      dedupeSkippedEmailCount:
        normalizeOutlookScanCount(response?.dedupeSkippedEmailCount) ??
        normalizeOutlookScanCount(
          outlookScanState.progress?.dedupeSkippedEmailCount
        ),
      source:
        String(response?.source || outlookScanState.progress?.source || "").trim(),
      timeframe: String(response?.timeframe || "").trim().toLowerCase(),
      scanDate: normalizeOutlookScanDateInput(response?.scanDate || scanDate),
      promptTruncated: !!response?.promptTruncated,
      errorMessage:
        response?.status === "success"
          ? ""
          : String(response?.message || "Outlook scan failed.").trim(),
    };
    if (response?.status === "success") {
      outlookScanState.progress = normalizeOutlookScanProgress(
        { active: false },
        outlookScanState.progress
      );
      outlookScanState.reportOpen = true;
      renderOutlookScanUi();
      toast(
        `${getOutlookScanSourceLabel(response?.source)} scan found ${
          derived.suggestions.length
        } suggestion${derived.suggestions.length === 1 ? "" : "s"}.`
      );
      return;
    }
    if (
      String(outlookScanState.progress?.stage || "").trim().toLowerCase() !==
        "error" ||
      outlookScanState.progress?.active
    ) {
      setOutlookScanProgress({
        active: false,
        stage: "error",
        message: response?.message || "Outlook scan failed.",
        source:
          String(response?.source || outlookScanState.progress?.source || "").trim(),
        timeframe: String(response?.timeframe || "").trim().toLowerCase(),
        scanDate: normalizeOutlookScanDateInput(response?.scanDate || scanDate),
        skippedEmails: derived.skipped.length,
        deliverablesInPeriod:
          normalizeOutlookScanCount(outlookScanState.progress?.deliverablesInPeriod) ??
          deliverablesInPeriod,
        threadsDetected:
          normalizeOutlookScanCount(response?.threadsDetected) ??
          normalizeOutlookScanCount(outlookScanState.progress?.threadsDetected),
        dedupedEmailCount:
          normalizeOutlookScanCount(response?.dedupedEmailCount) ??
          normalizeOutlookScanCount(outlookScanState.progress?.dedupedEmailCount),
        dedupeSkippedEmailCount:
          normalizeOutlookScanCount(response?.dedupeSkippedEmailCount) ??
          normalizeOutlookScanCount(
            outlookScanState.progress?.dedupeSkippedEmailCount
          ),
      });
    } else {
      outlookScanState.progress = normalizeOutlookScanProgress(
        { active: false },
        outlookScanState.progress
      );
    }
    outlookScanState.reportOpen = true;
    toast(
      response?.message || "Outlook scan failed."
    );
  } catch (e) {
    console.warn("Outlook scan failed:", e);
    const errorMessage = e?.message || "Outlook scan failed.";
    outlookScanState.lastResult = {
      status: "error",
      message: errorMessage,
      source: String(outlookScanState.progress?.source || "").trim(),
      timeframe: "",
      scanDate,
      suggestions: [],
      skippedMessages: [],
    };
    outlookScanState.suggestions = [];
    outlookScanState.skipped = [];
    outlookScanState.reportStats = {
      suggestionCount: 0,
      skippedCount: 0,
      deliverablesInPeriod:
        normalizeOutlookScanCount(outlookScanState.progress?.deliverablesInPeriod) ??
        deliverablesInPeriod,
      threadsDetected: normalizeOutlookScanCount(
        outlookScanState.progress?.threadsDetected
      ),
      dedupedEmailCount: normalizeOutlookScanCount(
        outlookScanState.progress?.dedupedEmailCount
      ),
      dedupeSkippedEmailCount: normalizeOutlookScanCount(
        outlookScanState.progress?.dedupeSkippedEmailCount
      ),
      source: String(outlookScanState.progress?.source || "").trim(),
      timeframe: "",
      scanDate,
      promptTruncated: false,
      errorMessage,
    };
    setOutlookScanProgress({
      active: false,
      stage: "error",
      message: errorMessage,
      source: String(outlookScanState.progress?.source || "").trim(),
      timeframe: "",
      scanDate,
      deliverablesInPeriod:
        normalizeOutlookScanCount(outlookScanState.progress?.deliverablesInPeriod) ??
        deliverablesInPeriod,
      threadsDetected: normalizeOutlookScanCount(
        outlookScanState.progress?.threadsDetected
      ),
      dedupedEmailCount: normalizeOutlookScanCount(
        outlookScanState.progress?.dedupedEmailCount
      ),
      dedupeSkippedEmailCount: normalizeOutlookScanCount(
        outlookScanState.progress?.dedupeSkippedEmailCount
      ),
    });
    outlookScanState.reportOpen = true;
    renderOutlookScanUi();
    toast(errorMessage);
  } finally {
    outlookScanState.busy = false;
    renderOutlookScanUi();
  }
}

function dismissOutlookScanSuggestion(suggestionKey) {
  const current = new Set(outlookScanState.dismissedKeys || []);
  if (suggestionKey) current.add(suggestionKey);
  outlookScanState.dismissedKeys = [...current];
  renderOutlookScanUi();
}

async function acceptOutlookScanSuggestion(suggestionKey) {
  const suggestion = (outlookScanState.suggestions || []).find(
    (entry) => entry?.key === suggestionKey
  );
  if (!suggestion || !db[suggestion.projectIndex]) return false;

  const snapshot = deepCloneJson(db[suggestion.projectIndex], null);
  const project = normalizeProject(db[suggestion.projectIndex]);
  const deliverable = createDeliverable({
    name: suggestion.deliverableName || "",
    due: suggestion.due || "",
    notes: suggestion.notes || "",
    tasks: suggestion.tasks || [],
    emailRefs: suggestion.emailRefs || [],
    emailRef: (suggestion.emailRefs || [])[0] || null,
  });
  project.deliverables.push(deliverable);
  syncProjectActiveDeliverables(project, {
    fallbackActiveId: deliverable.id,
  });
  db[suggestion.projectIndex] = project;

  const saved = await save({ silent: true });
  if (!saved) {
    if (snapshot) db[suggestion.projectIndex] = snapshot;
    render();
    toast("Could not save the new deliverable.");
    return false;
  }

  outlookScanState.suggestions = (outlookScanState.suggestions || []).filter(
    (entry) => entry?.key !== suggestionKey
  );
  outlookScanState.dismissedKeys = (outlookScanState.dismissedKeys || []).filter(
    (key) => key !== suggestionKey
  );
  render();
  renderOutlookScanUi();
  toast(
    `Added ${suggestion.deliverableName || "deliverable"} to ${
      suggestion.projectName || suggestion.projectId || "project"
    }.`
  );
  return true;
}

// ===================== CLOUD SYNC =====================

function updateCloudSyncState(patch = {}) {
  cloudSyncState = {
    ...cloudSyncState,
    ...patch,
  };
  if (patch.lastSyncedAt !== undefined) {
    cloudSyncState.lastSyncedAt = normalizeIsoTimestamp(patch.lastSyncedAt);
  }
  renderGoogleAuthUi();
}

async function updateLocalCloudSyncMetadata(
  patch = {},
  { persist = true } = {}
) {
  const current = ensureCloudSyncSettingsObject();
  userSettings.cloudSync = normalizeCloudSyncSettings({
    ...current,
    ...patch,
    lastSyncedAt:
      patch.lastSyncedAt !== undefined
        ? normalizeIsoTimestamp(patch.lastSyncedAt)
        : current.lastSyncedAt,
  });
  if (persist) {
    await persistUserSettingsLocally({
      skipCloud: true,
      saveTimestamp: false,
      silent: true,
    });
  }
  updateCloudSyncState({
    enabled: userSettings.cloudSync.enabled,
    firebaseUid: userSettings.cloudSync.firebaseUid,
    lastSyncedAt: userSettings.cloudSync.lastSyncedAt,
  });
  return userSettings.cloudSync;
}

function getCloudSyncDisplayModel() {
  const configured = cloudSyncState.configured === true;
  const lastSyncedAt =
    cloudSyncState.lastSyncedAt || ensureCloudSyncSettingsObject().lastSyncedAt;
  if (!configured) {
    return {
      status: "Cloud sync not configured",
      details:
        "Add FIREBASE_API_KEY, FIREBASE_AUTH_DOMAIN, FIREBASE_PROJECT_ID, and FIREBASE_APP_ID to .env to enable cross-device sync.",
      note:
        "Google sign-in works without Firebase, but cross-device sync stays disabled until Firebase is configured.",
      dot: "disabled",
    };
  }
  if (!googleAuthState.signedIn) {
    return {
      status: "Ready when you sign in",
      details:
        "Sign in with Google to sync settings, projects, notes, templates, checklists, and timesheets.",
      note:
        "This app stays local-first until a Google account is connected and Firebase sync is available.",
      dot: "idle",
    };
  }
  if (cloudSyncState.busy) {
    return {
      status: "Syncing...",
      details: cloudSyncState.message || "Connecting your Google account to Firestore.",
      note: "Sync is in progress. Keep the app open until the initial pull completes.",
      dot: "busy",
    };
  }
  if (cloudSyncState.error) {
    return {
      status: "Sync error",
      details: cloudSyncState.error,
      note:
        "Your local data is still available. Fix the Firebase configuration or connectivity issue and sign in again.",
      dot: "error",
    };
  }
  if (cloudSyncState.enabled) {
    return {
      status: "Cloud sync active",
      details: lastSyncedAt
        ? `Last sync: ${formatSyncTimestamp(lastSyncedAt)}`
        : "Cross-device sync is connected.",
      note:
        "Supported app data now syncs through Firestore. Local-only secrets and file paths stay on this device.",
      dot: "active",
    };
  }
  return {
    status: "Signed in locally",
    details: "Google is connected, but Firestore sync is not active yet.",
    note:
      "Your Google account is available locally, but cross-device sync has not been established.",
    dot: "idle",
  };
}

function renderCloudSyncUi() {
  const settingsStatus = document.getElementById("settings_cloudSyncStatus");
  const settingsDetails = document.getElementById("settings_cloudSyncDetails");
  const accountNote = document.getElementById("googleAccountSyncNote");
  const headerBtn = document.getElementById("headerGoogleAuthBtn");
  const headerDot = document.getElementById("headerGoogleAuthStatusDot");
  const headerPopover = document.getElementById("headerAccountPopover");
  const headerPopoverStatus = document.getElementById("headerAccountPopoverStatus");
  const headerPopoverNote = document.getElementById("headerAccountPopoverNote");
  const headerPopoverDot = document.getElementById("headerAccountPopoverSyncDot");
  const model = getCloudSyncDisplayModel();

  if (settingsStatus) {
    settingsStatus.textContent = model.status;
  }
  if (settingsDetails) {
    settingsDetails.textContent = model.details;
  }
  if (accountNote) {
    accountNote.textContent = model.note;
  }
  if (headerBtn) {
    headerBtn.dataset.syncStatus = model.dot;
  }
  if (headerDot) {
    headerDot.title = model.status;
  }
  if (headerPopover) {
    headerPopover.dataset.syncStatus = model.dot;
  }
  if (headerPopoverStatus) {
    headerPopoverStatus.textContent = model.status;
  }
  if (headerPopoverNote) {
    headerPopoverNote.textContent = model.note;
  }
  if (headerPopoverDot) {
    headerPopoverDot.title = model.status;
  }
}

async function loadCloudSyncConfig() {
  if (cloudSyncConfig) {
    updateCloudSyncState({
      configured: !!(
        cloudSyncConfig.apiKey &&
        cloudSyncConfig.authDomain &&
        cloudSyncConfig.projectId &&
        cloudSyncConfig.appId
      ),
      available: true,
    });
    return cloudSyncConfig;
  }
  if (!window.pywebview?.api?.get_cloud_sync_config) {
    updateCloudSyncState({
      configured: false,
      available: false,
      status: "local-only",
    });
    return null;
  }
  try {
    const response = await window.pywebview.api.get_cloud_sync_config();
    if (response?.status !== "success") {
      throw new Error(response?.message || "Failed to load cloud sync configuration.");
    }
    cloudSyncConfig = response.config || null;
    updateCloudSyncState({
      configured: response.enabled === true,
      available: true,
    });
    return cloudSyncConfig;
  } catch (e) {
    console.warn("Failed to load cloud sync config:", e);
    updateCloudSyncState({
      configured: false,
      available: false,
      error: String(e?.message || "Cloud sync config unavailable."),
    });
    return null;
  }
}

function loadExternalScript(src) {
  return new Promise((resolve, reject) => {
    const existing = document.querySelector(`script[data-cloud-sync-src="${src}"]`);
    if (existing?.dataset.loaded === "true") {
      resolve();
      return;
    }
    const script = existing || document.createElement("script");
    script.async = true;
    script.src = src;
    script.dataset.cloudSyncSrc = src;
    const timeoutId = window.setTimeout(() => {
      reject(new Error(`Timed out loading ${src}`));
    }, 15000);
    script.onload = () => {
      window.clearTimeout(timeoutId);
      script.dataset.loaded = "true";
      resolve();
    };
    script.onerror = () => {
      window.clearTimeout(timeoutId);
      reject(new Error(`Failed to load ${src}`));
    };
    if (!existing) {
      document.head.appendChild(script);
    }
  });
}

async function ensureFirebaseSdk() {
  if (window.firebase?.apps) {
    updateCloudSyncState({ sdkLoaded: true });
    return true;
  }
  if (!firebaseLoadPromise) {
    firebaseLoadPromise = (async () => {
      for (const src of FIREBASE_COMPAT_SCRIPT_URLS) {
        await loadExternalScript(src);
      }
      updateCloudSyncState({ sdkLoaded: true });
      return true;
    })().catch((error) => {
      firebaseLoadPromise = null;
      throw error;
    });
  }
  try {
    await firebaseLoadPromise;
    return true;
  } catch (e) {
    console.warn("Failed to load Firebase SDK:", e);
    updateCloudSyncState({
      sdkLoaded: false,
      error: String(e?.message || "Firebase SDK could not be loaded."),
    });
    return false;
  }
}

async function ensureFirebaseServices() {
  const config = await loadCloudSyncConfig();
  if (
    !config ||
    !config.apiKey ||
    !config.authDomain ||
    !config.projectId ||
    !config.appId
  ) {
    updateCloudSyncState({
      configured: false,
      enabled: false,
    });
    return null;
  }
  const sdkReady = await ensureFirebaseSdk();
  if (!sdkReady || !window.firebase?.initializeApp) {
    return null;
  }
  if (!firebaseAppInstance) {
    const existingApp =
      window.firebase.apps?.find((app) => app?.name === "acies-cloud-sync") ||
      null;
    firebaseAppInstance =
      existingApp || window.firebase.initializeApp(config, "acies-cloud-sync");
    firebaseAuthInstance = window.firebase.auth(firebaseAppInstance);
    firebaseFirestoreInstance = window.firebase.firestore(firebaseAppInstance);
    firebaseFirestoreInstance.settings({
      ignoreUndefinedProperties: true,
    });
  }
  return {
    app: firebaseAppInstance,
    auth: firebaseAuthInstance,
    db: firebaseFirestoreInstance,
  };
}

async function getGoogleSyncSessionFromBackend() {
  if (!window.pywebview?.api?.get_google_sync_session) {
    return {
      status: "error",
      signedIn: false,
      idToken: "",
      accessToken: "",
      firebaseReady: false,
      auth: { ...DEFAULT_GOOGLE_AUTH_STATE },
    };
  }
  const response = await window.pywebview.api.get_google_sync_session();
  return {
    status: response?.status || "error",
    signedIn: response?.signedIn === true,
    idToken: String(response?.idToken || "").trim(),
    accessToken: String(response?.accessToken || "").trim(),
    firebaseReady: response?.firebaseReady === true,
    auth: normalizeGoogleAuthState(response?.auth),
    message: response?.message || "",
  };
}

function getCloudSyncDocRefs(uid) {
  const userDoc = firebaseFirestoreInstance.collection("users").doc(uid);
  return {
    settings: userDoc.collection("settings").doc("app"),
    tasks: userDoc.collection("tasks").doc("main"),
    notes: userDoc.collection("notes").doc("main"),
    templates: userDoc.collection("templates").doc("main"),
    checklists: userDoc.collection("checklists").doc("main"),
    timesheets: userDoc.collection("timesheets"),
    timesheetsMeta: userDoc
      .collection("timesheets")
      .doc(CLOUD_SYNC_TIMESHEETS_META_DOC_ID),
  };
}

function clearCloudSyncSubscriptions() {
  cloudSyncUnsubscribers.forEach((unsubscribe) => {
    try {
      unsubscribe();
    } catch (e) {
      console.warn("Failed to unsubscribe cloud sync listener:", e);
    }
  });
  cloudSyncUnsubscribers = [];
}

async function signInToCloud(session = null) {
  const services = await ensureFirebaseServices();
  if (!services) {
    return null;
  }
  const syncSession = session || (await getGoogleSyncSessionFromBackend());
  if (!syncSession?.signedIn) {
    return null;
  }
  const idToken = String(syncSession.idToken || "").trim();
  const accessToken = String(syncSession.accessToken || "").trim();
  if (!idToken && !accessToken) {
    if (firebaseAuthInstance.currentUser) {
      return firebaseAuthInstance.currentUser;
    }
    throw new Error(
      "Google sign-in completed, but no Firebase-compatible token was returned."
    );
  }
  const credential = window.firebase.auth.GoogleAuthProvider.credential(
    idToken || null,
    accessToken || null
  );
  const userCredential = await firebaseAuthInstance.signInWithCredential(
    credential
  );
  return userCredential?.user || firebaseAuthInstance.currentUser;
}

function queueCloudStatePush(domain, delay = 900) {
  if (!cloudSyncState.enabled || isCloudSyncApplying()) return;
  if (domain === "timesheets") {
    if (cloudSyncTimesheetsPushTimer) {
      clearTimeout(cloudSyncTimesheetsPushTimer);
    }
    cloudSyncTimesheetsPushTimer = setTimeout(() => {
      pushUserState(["timesheets"]).catch((error) => {
        console.warn("Failed to push timesheets to cloud:", error);
        updateCloudSyncState({
          error: String(error?.message || "Timesheet sync failed."),
          status: "error",
        });
      });
    }, delay);
    return;
  }
  if (cloudSyncPushTimers[domain]) {
    clearTimeout(cloudSyncPushTimers[domain]);
  }
  cloudSyncPushTimers[domain] = setTimeout(() => {
    pushUserState([domain]).catch((error) => {
      console.warn(`Failed to push ${domain} to cloud:`, error);
      updateCloudSyncState({
        error: String(error?.message || `${domain} sync failed.`),
        status: "error",
      });
    });
  }, delay);
}

const schedulePullUserState = debounce(() => {
  pullUserState({ silent: true }).catch((error) => {
    console.warn("Failed to pull remote cloud state:", error);
    updateCloudSyncState({
      error: String(error?.message || "Cloud pull failed."),
      status: "error",
    });
  });
}, 700);

function subscribeToUserState(uid = cloudSyncState.firebaseUid) {
  if (!uid || !firebaseFirestoreInstance) return;
  clearCloudSyncSubscriptions();
  const refs = getCloudSyncDocRefs(uid);
  const listen = (ref) =>
    ref.onSnapshot(
      (snapshot) => {
        if (snapshot?.metadata?.hasPendingWrites) return;
        schedulePullUserState();
      },
      (error) => {
        console.warn("Cloud sync listener failed:", error);
        updateCloudSyncState({
          error: String(error?.message || "Cloud listener failed."),
          status: "error",
        });
      }
    );
  cloudSyncUnsubscribers = [
    listen(refs.settings),
    listen(refs.tasks),
    listen(refs.notes),
    listen(refs.templates),
    listen(refs.checklists),
    listen(refs.timesheetsMeta),
  ];
}

async function loadLocalSyncMetadata() {
  if (!window.pywebview?.api?.get_local_sync_metadata) return;
  try {
    const response = await window.pywebview.api.get_local_sync_metadata();
    if (response?.status !== "success") return;
    const files = response.files || {};
    localSyncTimestamps.settings = normalizeIsoTimestamp(
      files?.settings?.modified || localSyncTimestamps.settings
    );
    localSyncTimestamps.tasks = normalizeIsoTimestamp(
      files?.tasks?.modified || localSyncTimestamps.tasks
    );
    localSyncTimestamps.notes = normalizeIsoTimestamp(
      files?.notes?.modified || localSyncTimestamps.notes
    );
    localSyncTimestamps.templates = getLatestIsoTimestamp([
      files?.templates?.modified,
      templatesDb?.lastModified,
      localSyncTimestamps.templates,
    ]);
    localSyncTimestamps.checklists = getLatestIsoTimestamp([
      files?.checklists?.modified,
      checklistsDb?.lastModified,
      localSyncTimestamps.checklists,
    ]);
    localSyncTimestamps.timesheets = getLatestIsoTimestamp([
      files?.timesheets?.modified,
      timesheetDb?.lastModified,
      localSyncTimestamps.timesheets,
    ]);
  } catch (e) {
    console.warn("Failed to load local sync metadata:", e);
  }
}

async function createCloudSyncBackup(reason, metadata = {}) {
  if (!window.pywebview?.api?.create_cloud_sync_backup) return null;
  try {
    const response = await window.pywebview.api.create_cloud_sync_backup(
      reason,
      metadata
    );
    if (response?.status === "success") {
      return response;
    }
  } catch (e) {
    console.warn("Failed to create cloud sync backup:", e);
  }
  return null;
}

function getDefaultSyncableSettings() {
  return {
    userName: "",
    discipline: ["Electrical"],
    showSetupHelp: true,
    theme: "dark",
    lightingTemplates: [],
    autoPrimary: false,
    separateDeliverableCompletionGroups: true,
    groupDeliverablesByProject: false,
    defaultPmInitials: "",
    cleanDwgOptions: { ...DEFAULT_CLEAN_DWG_OPTIONS },
    publishDwgOptions: { ...DEFAULT_PUBLISH_DWG_OPTIONS },
    freezeLayerOptions: { ...DEFAULT_FREEZE_LAYER_OPTIONS },
    thawLayerOptions: { ...DEFAULT_THAW_LAYER_OPTIONS },
    workroomAutoSelectCadFiles: true,
    enableUnderConstructionTools: false,
  };
}

function sanitizeSettingsForCloud(settings = userSettings) {
  const source = settings && typeof settings === "object" ? settings : {};
  const updatedAt =
    normalizeIsoTimestamp(localSyncTimestamps.settings) || new Date().toISOString();
  return {
    ...getDefaultSyncableSettings(),
    userName: String(source.userName || "").trim(),
    discipline: normalizeDisciplineList(source.discipline),
    showSetupHelp: source.showSetupHelp !== false,
    theme: source.theme === "light" ? "light" : "dark",
    lightingTemplates: Array.isArray(source.lightingTemplates)
      ? deepCloneJson(source.lightingTemplates, [])
      : [],
    autoPrimary: source.autoPrimary === true,
    separateDeliverableCompletionGroups:
      source.separateDeliverableCompletionGroups !== false,
    groupDeliverablesByProject: source.groupDeliverablesByProject === true,
    defaultPmInitials: String(source.defaultPmInitials || "")
      .trim()
      .toUpperCase(),
    cleanDwgOptions: {
      ...DEFAULT_CLEAN_DWG_OPTIONS,
      ...(source.cleanDwgOptions || {}),
    },
    publishDwgOptions: {
      ...DEFAULT_PUBLISH_DWG_OPTIONS,
      ...(source.publishDwgOptions || {}),
    },
    freezeLayerOptions: {
      ...DEFAULT_FREEZE_LAYER_OPTIONS,
      ...(source.freezeLayerOptions || {}),
    },
    thawLayerOptions: {
      ...DEFAULT_THAW_LAYER_OPTIONS,
      ...(source.thawLayerOptions || {}),
    },
    workroomAutoSelectCadFiles: source.workroomAutoSelectCadFiles !== false,
    enableUnderConstructionTools: source.enableUnderConstructionTools === true,
    updatedAt,
  };
}

function normalizeCloudSettingsDoc(raw = {}) {
  const source = raw && typeof raw === "object" ? raw : {};
  const defaults = getDefaultSyncableSettings();
  return {
    ...defaults,
    userName: String(source.userName || "").trim(),
    discipline: normalizeDisciplineList(source.discipline),
    showSetupHelp: source.showSetupHelp !== false,
    theme: source.theme === "light" ? "light" : "dark",
    lightingTemplates: Array.isArray(source.lightingTemplates)
      ? deepCloneJson(source.lightingTemplates, [])
      : [],
    autoPrimary: source.autoPrimary === true,
    separateDeliverableCompletionGroups:
      source.separateDeliverableCompletionGroups !== false,
    groupDeliverablesByProject: source.groupDeliverablesByProject === true,
    defaultPmInitials: String(source.defaultPmInitials || "")
      .trim()
      .toUpperCase(),
    cleanDwgOptions: {
      ...DEFAULT_CLEAN_DWG_OPTIONS,
      ...(source.cleanDwgOptions || {}),
    },
    publishDwgOptions: {
      ...DEFAULT_PUBLISH_DWG_OPTIONS,
      ...(source.publishDwgOptions || {}),
    },
    freezeLayerOptions: {
      ...DEFAULT_FREEZE_LAYER_OPTIONS,
      ...(source.freezeLayerOptions || {}),
    },
    thawLayerOptions: {
      ...DEFAULT_THAW_LAYER_OPTIONS,
      ...(source.thawLayerOptions || {}),
    },
    workroomAutoSelectCadFiles: source.workroomAutoSelectCadFiles !== false,
    enableUnderConstructionTools: source.enableUnderConstructionTools === true,
    updatedAt:
      normalizeIsoTimestamp(source.updatedAt) ||
      normalizeIsoTimestamp(source.lastModified),
  };
}

function hasMeaningfulSettingsState(doc) {
  if (!doc) return false;
  const comparable = deepCloneJson(doc, {}) || {};
  delete comparable.updatedAt;
  return stableStringify(comparable) !== stableStringify(getDefaultSyncableSettings());
}

function sanitizeLinkForCloud(link) {
  const normalized = normalizeRef(link);
  if (!normalized || isLocalOnlyLink(normalized)) return null;
  return normalized;
}

function getLinkKey(link) {
  const normalized = normalizeRef(link);
  if (!normalized) return "";
  return `${String(normalized.raw || normalized.url || "")
    .trim()
    .toLowerCase()}|${String(normalized.label || "")
    .trim()
    .toLowerCase()}`;
}

function mergeCloudAndLocalLinks(remoteRefs = [], localRefs = []) {
  const merged = [];
  const seen = new Set();
  [...(Array.isArray(remoteRefs) ? remoteRefs : [])]
    .map((ref) => normalizeRef(ref))
    .filter(Boolean)
    .forEach((ref) => {
      const key = getLinkKey(ref);
      if (!key || seen.has(key)) return;
      seen.add(key);
      merged.push(ref);
    });
  [...(Array.isArray(localRefs) ? localRefs : [])]
    .map((ref) => normalizeRef(ref))
    .filter((ref) => ref && isLocalOnlyLink(ref))
    .forEach((ref) => {
      const key = getLinkKey(ref);
      if (!key || seen.has(key)) return;
      seen.add(key);
      merged.push(ref);
    });
  return merged;
}

function sanitizeEmailRefForCloud(ref) {
  const normalized = normalizeEmailRef(ref);
  if (!normalized) return null;
  if (
    String(normalized.source || "").toLowerCase() === "file" ||
    isLikelyLocalPath(normalized.raw) ||
    isLikelyLocalPath(normalized.url)
  ) {
    return null;
  }
  return normalized;
}

function sanitizeAttachmentForCloud(attachment = {}) {
  const normalized = normalizeAttachmentEntry(attachment);
  if (!normalized) return null;
  if (normalized.type === "email") {
    const emailRef = sanitizeEmailRefForCloud(normalized.emailRef);
    if (!emailRef) return null;
    return {
      id: normalized.id,
      type: "email",
      description: String(normalized.description || "").trim(),
      emailRef,
    };
  }
  if (normalized.type !== "url") {
    return null;
  }
  return {
    id: normalized.id,
    type: "url",
    description: String(normalized.description || "").trim(),
    target: normalized.target,
  };
}

function mergeCloudAndLocalAttachments(remoteAttachments = [], localAttachments = []) {
  const merged = [];
  const seen = new Map();
  [
    ...(Array.isArray(remoteAttachments) ? remoteAttachments : []),
    ...(Array.isArray(localAttachments) ? localAttachments : []),
  ]
    .map((attachment) => normalizeAttachmentEntry(attachment))
    .filter(Boolean)
    .forEach((attachment) => {
      const key = getAttachmentEntryKey(attachment);
      if (!key) return;
      if (seen.has(key)) {
        const existing = seen.get(key);
        if (!existing.description && attachment.description) {
          existing.description = attachment.description;
        }
        return;
      }
      seen.set(key, attachment);
      merged.push(attachment);
    });
  return merged;
}

function mergeCloudAndLocalEmailRefs(remoteRefs = [], localRefs = []) {
  const merged = [];
  const seen = new Set();
  [
    ...normalizeEmailRefs(remoteRefs, null),
    ...normalizeEmailRefs(localRefs, null).filter(
      (ref) =>
        ref &&
        (String(ref.source || "").toLowerCase() === "file" ||
          isLikelyLocalPath(ref.raw) ||
          isLikelyLocalPath(ref.url))
    ),
  ].forEach((ref) => {
    const normalized = normalizeEmailRef(ref);
    const key = normalizeEmailRefKey(normalized);
    if (!normalized || !key || seen.has(key)) return;
    seen.add(key);
    merged.push(normalized);
  });
  return merged;
}

function sanitizeDeliverableForCloud(deliverable = {}) {
  const source =
    deliverable && typeof deliverable === "object"
      ? deepCloneJson(deliverable, {}) || {}
      : {};
  const normalized = normalizeDeliverable(source);
  const attachments = normalizeAttachments(normalized.attachments)
    .map((attachment) => sanitizeAttachmentForCloud(attachment))
    .filter(Boolean);
  const emailRefs = buildLegacyEmailRefsFromAttachments(attachments);
  return {
    ...source,
    ...normalized,
    attachments,
    links: buildLegacyLinksFromAttachments(attachments),
    emailRefs,
    emailRef: emailRefs[0] || null,
    tasks: (Array.isArray(normalized.tasks) ? normalized.tasks : []).map((task) => {
      const taskAttachments = normalizeAttachments(task.attachments)
        .map((attachment) => sanitizeAttachmentForCloud(attachment))
        .filter(Boolean);
      return {
        ...task,
        attachments: taskAttachments,
        links: buildLegacyLinksFromAttachments(taskAttachments),
        emailRefs: buildLegacyEmailRefsFromAttachments(taskAttachments),
      };
    }),
    noteItems: normalizeDeliverableNoteItems(
      normalized.noteItems,
      normalized.notes || ""
    ).map((noteItem) => {
      const noteAttachments = normalizeAttachments(noteItem.attachments)
        .map((attachment) => sanitizeAttachmentForCloud(attachment))
        .filter(Boolean);
      return {
        ...noteItem,
        attachments: noteAttachments,
        links: buildLegacyLinksFromAttachments(noteAttachments),
        emailRefs: buildLegacyEmailRefsFromAttachments(noteAttachments),
      };
    }),
  };
}

function getCloudDeliverableKey(deliverable, index = 0) {
  const id = String(deliverable?.id || "").trim().toLowerCase();
  if (id) return `id:${id}`;
  const name = String(deliverable?.name || "").trim().toLowerCase();
  const due = String(deliverable?.due || "").trim().toLowerCase();
  if (name || due) return `name:${name}|due:${due}`;
  return `index:${index}`;
}

function mergeLocalDeliverableFields(remoteDeliverable, localDeliverable) {
  const merged = {
    ...(deepCloneJson(localDeliverable, {}) || {}),
    ...(deepCloneJson(remoteDeliverable, {}) || {}),
  };
  merged.attachments = mergeCloudAndLocalAttachments(
    normalizeAttachments(remoteDeliverable?.attachments, {
      legacyLinks: remoteDeliverable?.links,
      legacyEmailRefs: remoteDeliverable?.emailRefs,
      legacyEmailRef: remoteDeliverable?.emailRef,
    }),
    normalizeAttachments(localDeliverable?.attachments, {
      legacyLinks: localDeliverable?.links,
      legacyEmailRefs: localDeliverable?.emailRefs,
      legacyEmailRef: localDeliverable?.emailRef,
    })
  );
  const emailRefs = mergeCloudAndLocalEmailRefs(
    remoteDeliverable?.emailRefs,
    localDeliverable?.emailRefs
  );
  merged.emailRefs = emailRefs;
  merged.emailRef = emailRefs[0] || null;
  return merged;
}

function sanitizeProjectForCloud(project = {}) {
  const source =
    project && typeof project === "object"
      ? deepCloneJson(project, {}) || {}
      : {};
  const normalized = normalizeProject(source);
  const attachments = normalizeAttachments(normalized.attachments)
    .map((attachment) => sanitizeAttachmentForCloud(attachment))
    .filter(Boolean);
  return {
    ...source,
    ...normalized,
    path: "",
    localProjectPath: "",
    workroomRootPath: "",
    attachments,
    refs: (Array.isArray(source.refs) ? source.refs : [])
      .map((ref) => sanitizeLinkForCloud(ref))
      .filter(Boolean),
    links: buildLegacyLinksFromAttachments(attachments),
    deliverables: getProjectDeliverables(normalized).map((deliverable) =>
      sanitizeDeliverableForCloud(deliverable)
    ),
    lightingSchedule: normalized.lightingSchedule
      ? {
          ...deepCloneJson(normalized.lightingSchedule, {}),
          targetDwgPath: "",
        }
      : createDefaultLightingSchedule(),
    title24: normalized.title24
      ? {
          ...deepCloneJson(normalized.title24, {}),
          roomAreas: {
            ...(deepCloneJson(normalized.title24.roomAreas, {}) || {}),
            sourcePath: "",
          },
        }
      : createDefaultTitle24(),
  };
}

function getCloudProjectKey(project, index = 0) {
  const id = String(project?.id || "").trim().toLowerCase();
  if (id) return `id:${id}`;
  const name = String(project?.name || "").trim().toLowerCase();
  const nick = String(project?.nick || "").trim().toLowerCase();
  if (name || nick) return `name:${name}|nick:${nick}`;
  return `index:${index}`;
}

function mergeLocalProjectFields(remoteProject, localProject) {
  const merged = {
    ...(deepCloneJson(localProject, {}) || {}),
    ...(deepCloneJson(remoteProject, {}) || {}),
  };
  merged.path = String(localProject?.path || merged.path || "").trim();
  merged.localProjectPath = String(
    localProject?.localProjectPath || merged.localProjectPath || ""
  ).trim();
  merged.workroomRootPath = String(
    localProject?.workroomRootPath || merged.workroomRootPath || ""
  ).trim();
  merged.refs = mergeCloudAndLocalLinks(remoteProject?.refs, localProject?.refs);
  merged.attachments = mergeCloudAndLocalAttachments(
    normalizeAttachments(remoteProject?.attachments, {
      legacyLinks: remoteProject?.links,
    }),
    normalizeAttachments(localProject?.attachments, {
      legacyLinks: localProject?.links,
    })
  );

  const localDeliverables = Array.isArray(localProject?.deliverables)
    ? localProject.deliverables
    : [];
  const localDeliverableMap = new Map(
    localDeliverables.map((deliverable, index) => [
      getCloudDeliverableKey(deliverable, index),
      deliverable,
    ])
  );
  merged.deliverables = getProjectDeliverables(remoteProject).map(
    (deliverable, index) =>
      mergeLocalDeliverableFields(
        deliverable,
        localDeliverableMap.get(getCloudDeliverableKey(deliverable, index))
      )
  );

  if (merged.lightingSchedule) {
    merged.lightingSchedule.targetDwgPath = String(
      localProject?.lightingSchedule?.targetDwgPath ||
        merged.lightingSchedule.targetDwgPath ||
        ""
    ).trim();
  }
  if (merged.title24?.roomAreas) {
    merged.title24.roomAreas.sourcePath = normalizeTitle24Text(
      localProject?.title24?.roomAreas?.sourcePath ||
        merged.title24.roomAreas.sourcePath ||
        ""
    );
  }
  return normalizeProject(merged);
}

function restoreLocalOnlyProjectFields(remoteProjects, localProjects = db) {
  const localMap = new Map(
    (Array.isArray(localProjects) ? localProjects : []).map((project, index) => [
      getCloudProjectKey(project, index),
      project,
    ])
  );
  return (Array.isArray(remoteProjects) ? remoteProjects : [])
    .map((project, index) =>
      mergeLocalProjectFields(
        project,
        localMap.get(getCloudProjectKey(project, index))
      )
    )
    .filter(Boolean);
}

function buildTasksCloudDoc() {
  return {
    projects: deepCloneJson(db, []).map((project) => sanitizeProjectForCloud(project)),
    updatedAt:
      normalizeIsoTimestamp(localSyncTimestamps.tasks) || new Date().toISOString(),
  };
}

function normalizeCloudTasksDoc(raw = {}) {
  const source = Array.isArray(raw)
    ? { projects: raw }
    : raw && typeof raw === "object"
      ? raw
      : {};
  const migrated = migrateProjects(
    Array.isArray(source.projects) ? source.projects : []
  );
  return {
    projects: migrated.data,
    updatedAt:
      normalizeIsoTimestamp(source.updatedAt) ||
      normalizeIsoTimestamp(source.lastModified),
  };
}

function hasMeaningfulTasksState(doc) {
  return Array.isArray(doc?.projects) && doc.projects.length > 0;
}

function buildNotesCloudDoc() {
  return {
    tabs: [...(Array.isArray(noteTabs) && noteTabs.length ? noteTabs : ["General"])],
    general: deepCloneJson(notesDb, {}),
    updatedAt:
      normalizeIsoTimestamp(localSyncTimestamps.notes) || new Date().toISOString(),
  };
}

function normalizeCloudNotesDoc(raw = {}) {
  const source = raw && typeof raw === "object" ? raw : {};
  const tabs =
    Array.isArray(source.tabs) && source.tabs.length
      ? source.tabs.map((tab) => String(tab || "").trim()).filter(Boolean)
      : ["General"];
  const general =
    source.general && typeof source.general === "object" && !Array.isArray(source.general)
      ? deepCloneJson(source.general, {})
      : {};
  return {
    tabs: tabs.length ? tabs : ["General"],
    general,
    updatedAt:
      normalizeIsoTimestamp(source.updatedAt) ||
      normalizeIsoTimestamp(source.lastModified),
  };
}

function hasMeaningfulNotesState(doc) {
  if (!doc) return false;
  const values = Object.values(doc.general || {});
  const hasText = values.some((value) => String(value || "").trim());
  const tabs = Array.isArray(doc.tabs) ? doc.tabs : [];
  return hasText || tabs.some((tab) => String(tab || "").trim() && tab !== "General");
}

function buildTemplateSourceName(template) {
  return basename(template?.sourcePath || template?.sourceName || "");
}

function sanitizeTemplateForCloud(template = {}) {
  const source =
    template && typeof template === "object"
      ? deepCloneJson(template, {}) || {}
      : {};
  const sourceName = buildTemplateSourceName(source);
  return {
    ...source,
    sourcePath: "",
    sourceName,
    cloudNeedsRelink: !source.isDefault && !!sourceName,
  };
}

function mergeLocalTemplateFields(remoteTemplate, localTemplate) {
  const merged = {
    ...(deepCloneJson(remoteTemplate, {}) || {}),
  };
  const localSourcePath = String(localTemplate?.sourcePath || "").trim();
  if (!String(merged.sourcePath || "").trim() && localSourcePath) {
    merged.sourcePath = localSourcePath;
    merged.cloudNeedsRelink = false;
  }
  if (!String(merged.sourceName || "").trim()) {
    merged.sourceName = buildTemplateSourceName(localTemplate) || "";
  }
  return merged;
}

function mergeLocalTemplatePaths(remoteDoc, localDoc = templatesDb) {
  const localMap = new Map(
    (Array.isArray(localDoc?.templates) ? localDoc.templates : []).map((template) => [
      String(template?.id || "").trim(),
      template,
    ])
  );
  return {
    templates: (Array.isArray(remoteDoc?.templates) ? remoteDoc.templates : []).map(
      (template) =>
        mergeLocalTemplateFields(
          template,
          localMap.get(String(template?.id || "").trim())
        )
    ),
    defaultTemplatesInstalled: remoteDoc?.defaultTemplatesInstalled === true,
    lastModified: normalizeIsoTimestamp(
      remoteDoc?.lastModified || remoteDoc?.updatedAt
    ),
    updatedAt: normalizeIsoTimestamp(remoteDoc?.updatedAt),
  };
}

function buildTemplatesCloudDoc() {
  const updatedAt =
    normalizeIsoTimestamp(localSyncTimestamps.templates) ||
    normalizeIsoTimestamp(templatesDb?.lastModified) ||
    new Date().toISOString();
  return {
    templates: (Array.isArray(templatesDb?.templates) ? templatesDb.templates : []).map(
      (template) => sanitizeTemplateForCloud(template)
    ),
    defaultTemplatesInstalled: templatesDb?.defaultTemplatesInstalled === true,
    lastModified: updatedAt,
    updatedAt,
  };
}

function normalizeCloudTemplatesDoc(raw = {}) {
  const source = raw && typeof raw === "object" ? raw : {};
  return {
    templates: (Array.isArray(source.templates) ? source.templates : []).map((template) => ({
      ...(deepCloneJson(template, {}) || {}),
      sourcePath: String(template?.sourcePath || "").trim(),
      sourceName: String(template?.sourceName || "").trim(),
      cloudNeedsRelink:
        template?.cloudNeedsRelink === true ||
        (!template?.isDefault &&
          !String(template?.sourcePath || "").trim() &&
          !!String(template?.sourceName || "").trim()),
    })),
    defaultTemplatesInstalled: source.defaultTemplatesInstalled === true,
    lastModified:
      normalizeIsoTimestamp(source.lastModified) ||
      normalizeIsoTimestamp(source.updatedAt),
    updatedAt:
      normalizeIsoTimestamp(source.updatedAt) ||
      normalizeIsoTimestamp(source.lastModified),
  };
}

function hasMeaningfulTemplatesState(doc) {
  return Array.isArray(doc?.templates) && doc.templates.length > 0;
}

function buildChecklistsCloudDoc() {
  const updatedAt =
    normalizeIsoTimestamp(localSyncTimestamps.checklists) ||
    normalizeIsoTimestamp(checklistsDb?.lastModified) ||
    new Date().toISOString();
  return {
    ...deepCloneJson(checklistsDb, {
      checklists: [],
      templateOverrides: {},
      lastModified: null,
    }),
    lastModified: updatedAt,
    updatedAt,
  };
}

function normalizeCloudChecklistsDoc(raw = {}) {
  const source = raw && typeof raw === "object" ? raw : {};
  const normalized = normalizeChecklistsDb(source || {});
  return {
    ...normalized.data,
    updatedAt:
      normalizeIsoTimestamp(source.updatedAt) ||
      normalizeIsoTimestamp(source.lastModified) ||
      normalizeIsoTimestamp(normalized.data.lastModified),
  };
}

function hasMeaningfulChecklistsState(doc) {
  return (
    (Array.isArray(doc?.checklists) && doc.checklists.length > 0) ||
    (doc?.templateOverrides &&
      Object.keys(doc.templateOverrides).length > 0)
  );
}

function buildTimesheetsCloudState() {
  const updatedAt =
    normalizeIsoTimestamp(localSyncTimestamps.timesheets) ||
    normalizeIsoTimestamp(timesheetDb?.lastModified) ||
    new Date().toISOString();
  const weeks =
    timesheetDb?.weeks && typeof timesheetDb.weeks === "object"
      ? deepCloneJson(timesheetDb.weeks, {})
      : {};
  const knownWeeks = Object.keys(weeks || {}).filter(Boolean).sort();
  return {
    weeks,
    knownWeeks,
    updatedAt,
  };
}

function normalizeRemoteTimesheetsState(metaDoc = {}, docs = []) {
  const weeks = {};
  const knownWeeks = [];
  let updatedAt =
    normalizeIsoTimestamp(metaDoc?.updatedAt) ||
    normalizeIsoTimestamp(metaDoc?.lastModified);
  (Array.isArray(docs) ? docs : []).forEach((doc) => {
    const docId = String(doc?.id || "").trim();
    if (!docId || docId === CLOUD_SYNC_TIMESHEETS_META_DOC_ID) return;
    const payload = doc?.data && typeof doc.data === "object" ? doc.data : {};
    weeks[docId] = deepCloneJson(payload.data, {}) || {};
    knownWeeks.push(docId);
    if (isIsoAfter(payload.updatedAt, updatedAt)) {
      updatedAt = normalizeIsoTimestamp(payload.updatedAt);
    }
  });
  const metaKnownWeeks = Array.isArray(metaDoc?.knownWeeks)
    ? metaDoc.knownWeeks.map((weekKey) => String(weekKey || "").trim()).filter(Boolean)
    : [];
  return {
    weeks,
    knownWeeks: [...new Set(metaKnownWeeks.length ? metaKnownWeeks : knownWeeks)].sort(),
    updatedAt,
  };
}

function hasMeaningfulTimesheetsState(state) {
  return Object.keys(state?.weeks || {}).length > 0;
}

function getLocalCloudDoc(domain) {
  if (domain === "settings") return sanitizeSettingsForCloud(userSettings);
  if (domain === "tasks") return buildTasksCloudDoc();
  if (domain === "notes") return buildNotesCloudDoc();
  if (domain === "templates") return buildTemplatesCloudDoc();
  if (domain === "checklists") return buildChecklistsCloudDoc();
  if (domain === "timesheets") return buildTimesheetsCloudState();
  return null;
}

function hasMeaningfulCloudDoc(domain, doc) {
  if (domain === "settings") return hasMeaningfulSettingsState(doc);
  if (domain === "tasks") return hasMeaningfulTasksState(doc);
  if (domain === "notes") return hasMeaningfulNotesState(doc);
  if (domain === "templates") return hasMeaningfulTemplatesState(doc);
  if (domain === "checklists") return hasMeaningfulChecklistsState(doc);
  if (domain === "timesheets") return hasMeaningfulTimesheetsState(doc);
  return false;
}

function buildComparableCloudDoc(domain, doc = getLocalCloudDoc(domain)) {
  const comparable = deepCloneJson(doc, null);
  if (!comparable || typeof comparable !== "object") return comparable;
  delete comparable.updatedAt;
  if (domain === "templates" || domain === "checklists") {
    delete comparable.lastModified;
  }
  return comparable;
}

function getCloudComparableFingerprint(domain, doc = getLocalCloudDoc(domain)) {
  return stableStringify(buildComparableCloudDoc(domain, doc));
}

function syncCloudComparableFingerprint(domain, doc = getLocalCloudDoc(domain)) {
  const fingerprint = getCloudComparableFingerprint(domain, doc);
  lastCloudComparableFingerprints[domain] = fingerprint;
  return fingerprint;
}

function syncAllCloudComparableFingerprints() {
  [
    "settings",
    "tasks",
    "notes",
    "templates",
    "checklists",
    "timesheets",
  ].forEach((domain) => {
    syncCloudComparableFingerprint(domain);
  });
}

async function prepareLocalStateForUserSwitch(nextUid) {
  const currentMetadata = ensureCloudSyncSettingsObject();
  const previousUid = String(currentMetadata.firebaseUid || "").trim();
  if (!previousUid || previousUid === nextUid) return false;

  await createCloudSyncBackup("cloud-user-switch", {
    previousFirebaseUid: previousUid,
    nextFirebaseUid: nextUid,
  });

  const resetAt = new Date().toISOString();
  userSettings = {
    ...userSettings,
    ...getDefaultSyncableSettings(),
    googleAuth: userSettings.googleAuth,
    cloudSync: {
      ...currentMetadata,
      enabled: false,
      migrationCompleted: false,
      lastSyncedAt: "",
    },
  };
  db = [];
  notesDb = {};
  noteTabs = ["General"];
  activeNoteTab = "General";
  activeNotebookType = "note";
  notesSearchQuery = "";
  checklistSearchQuery = "";
  timesheetDb = { weeks: {}, lastModified: null };
  templatesDb = {
    templates: [],
    defaultTemplatesInstalled: false,
    lastModified: null,
  };
  checklistsDb = {
    checklists: [],
    templateOverrides: {},
    lastModified: null,
  };
  activeChecklistTabId = null;
  cloudSyncRemoteTimesheetMeta = {
    updatedAt: "",
    knownWeeks: [],
  };

  await persistUserSettingsLocally({
    skipCloud: true,
    timestamp: resetAt,
    silent: true,
  });
  await save({ skipCloud: true, timestamp: resetAt, silent: true });
  await saveNotes({ skipCloud: true, timestamp: resetAt, silent: true });
  await saveTimesheets({ skipCloud: true, timestamp: resetAt, silent: true });
  await saveTemplates({ skipCloud: true, timestamp: resetAt, silent: true });
  templatesDb = (await loadTemplates()) || templatesDb;
  await saveChecklists({ skipCloud: true, timestamp: resetAt, silent: true });
  syncAllCloudComparableFingerprints();
  renderNoteTabs();
  renderChecklistTabs();
  renderChecklistItems();
  renderTextsView();
  renderTemplates();
  renderTimesheets();
  render();
  return true;
}

async function pushCloudDomain(domain) {
  if (!cloudSyncState.enabled || !firebaseFirestoreInstance || !cloudSyncState.firebaseUid) {
    return "";
  }
  const refs = getCloudSyncDocRefs(cloudSyncState.firebaseUid);
  if (domain === "timesheets") {
    return pushCloudTimesheetsState();
  }
  const payload = getLocalCloudDoc(domain);
  if (!payload) return "";
  const updatedAt = normalizeIsoTimestamp(payload.updatedAt) || new Date().toISOString();
  const ref = refs[domain];
  if (!ref) return "";
  await ref.set(payload, { merge: false });
  return updatedAt;
}

async function pushCloudTimesheetsState() {
  if (!cloudSyncState.enabled || !firebaseFirestoreInstance || !cloudSyncState.firebaseUid) {
    return "";
  }
  const refs = getCloudSyncDocRefs(cloudSyncState.firebaseUid);
  const state = buildTimesheetsCloudState();
  const upserts = state.knownWeeks.map((weekKey) =>
    refs.timesheets.doc(weekKey).set(
      {
        weekKey,
        data: deepCloneJson(state.weeks[weekKey], {}),
        updatedAt: state.updatedAt,
      },
      { merge: false }
    )
  );
  const previousKnownWeeks = Array.isArray(cloudSyncRemoteTimesheetMeta.knownWeeks)
    ? cloudSyncRemoteTimesheetMeta.knownWeeks
    : [];
  const removals = previousKnownWeeks
    .filter((weekKey) => !state.knownWeeks.includes(weekKey))
    .map((weekKey) => refs.timesheets.doc(weekKey).delete());
  await Promise.all([
    ...upserts,
    ...removals,
    refs.timesheetsMeta.set(
      {
        updatedAt: state.updatedAt,
        knownWeeks: state.knownWeeks,
      },
      { merge: false }
    ),
  ]);
  cloudSyncRemoteTimesheetMeta = {
    updatedAt: state.updatedAt,
    knownWeeks: state.knownWeeks,
  };
  return state.updatedAt;
}

async function pushTimesheetWeek(weekKey) {
  if (!weekKey) return "";
  return pushCloudTimesheetsState();
}

async function pushUserState(domains = null) {
  if (!cloudSyncState.enabled || isCloudSyncApplying()) return "";
  const requested =
    Array.isArray(domains) && domains.length
      ? domains
      : ["settings", "tasks", "notes", "templates", "checklists", "timesheets"];
  const uniqueDomains = [...new Set(requested)];
  updateCloudSyncState({
    busy: true,
    status: "syncing",
    message: "Uploading local changes...",
    error: "",
  });
  try {
    const timestamps = [];
    for (const domain of uniqueDomains) {
      const updatedAt = await pushCloudDomain(domain);
      if (updatedAt) timestamps.push(updatedAt);
    }
    const lastSyncedAt = new Date().toISOString();
    await updateLocalCloudSyncMetadata(
      {
        enabled: true,
        firebaseUid: cloudSyncState.firebaseUid,
        migrationCompleted: true,
        lastSyncedAt,
      },
      { persist: true }
    );
    updateCloudSyncState({
      busy: false,
      enabled: true,
      status: "synced",
      message: lastSyncedAt
        ? `Last sync ${formatSyncTimestamp(lastSyncedAt)}`
        : "Cloud sync ready",
      error: "",
      lastSyncedAt,
    });
    return lastSyncedAt;
  } catch (e) {
    updateCloudSyncState({
      busy: false,
      status: "error",
      error: String(e?.message || "Cloud push failed."),
    });
    throw e;
  }
}

async function fetchRemoteUserState(uid = cloudSyncState.firebaseUid) {
  if (!uid || !firebaseFirestoreInstance) {
    return {
      settings: null,
      tasks: null,
      notes: null,
      templates: null,
      checklists: null,
      timesheets: normalizeRemoteTimesheetsState({}, []),
    };
  }
  const refs = getCloudSyncDocRefs(uid);
  const [
    settingsSnap,
    tasksSnap,
    notesSnap,
    templatesSnap,
    checklistsSnap,
    timesheetsSnap,
  ] = await Promise.all([
    refs.settings.get(),
    refs.tasks.get(),
    refs.notes.get(),
    refs.templates.get(),
    refs.checklists.get(),
    refs.timesheets.get(),
  ]);

  let timesheetsMeta = {};
  const timesheetDocs = [];
  timesheetsSnap.forEach((doc) => {
    if (doc.id === CLOUD_SYNC_TIMESHEETS_META_DOC_ID) {
      timesheetsMeta = doc.data() || {};
      return;
    }
    timesheetDocs.push({
      id: doc.id,
      data: doc.data() || {},
    });
  });

  return {
    settings: settingsSnap.exists ? normalizeCloudSettingsDoc(settingsSnap.data()) : null,
    tasks: tasksSnap.exists ? normalizeCloudTasksDoc(tasksSnap.data()) : null,
    notes: notesSnap.exists ? normalizeCloudNotesDoc(notesSnap.data()) : null,
    templates: templatesSnap.exists
      ? normalizeCloudTemplatesDoc(templatesSnap.data())
      : null,
    checklists: checklistsSnap.exists
      ? normalizeCloudChecklistsDoc(checklistsSnap.data())
      : null,
    timesheets: normalizeRemoteTimesheetsState(timesheetsMeta, timesheetDocs),
  };
}

async function applyRemoteCloudDoc(domain, remoteDoc) {
  beginCloudSyncApply();
  try {
    if (domain === "settings") {
      const cloudSettings = normalizeCloudSettingsDoc(remoteDoc);
      const currentCloudSync = ensureCloudSyncSettingsObject();
      userSettings = {
        ...userSettings,
        ...cloudSettings,
        apiKey: userSettings.apiKey,
        autocadPath: userSettings.autocadPath,
        googleAuth: userSettings.googleAuth,
        cloudSync: currentCloudSync,
      };
      touchLocalSyncTimestamp("settings", cloudSettings.updatedAt);
      await persistUserSettingsLocally({
        skipCloud: true,
        saveTimestamp: false,
        silent: true,
      });
      syncProjectViewPreferencesFromSettings();
      applyTheme(userSettings.theme);
      syncUnderConstructionToolsAvailability();
      render();
      return cloudSettings.updatedAt;
    }
    if (domain === "tasks") {
      const cloudTasks = normalizeCloudTasksDoc(remoteDoc);
      db = restoreLocalOnlyProjectFields(cloudTasks.projects, db);
      touchLocalSyncTimestamp("tasks", cloudTasks.updatedAt);
      await save({
        skipCloud: true,
        saveTimestamp: false,
        silent: true,
      });
      render();
      return cloudTasks.updatedAt;
    }
    if (domain === "notes") {
      const cloudNotes = normalizeCloudNotesDoc(remoteDoc);
      noteTabs = [...cloudNotes.tabs];
      notesDb = deepCloneJson(cloudNotes.general, {});
      activeNoteTab = noteTabs.includes(activeNoteTab) ? activeNoteTab : noteTabs[0];
      ensureActiveNotebookSelection(activeNotebookType);
      touchLocalSyncTimestamp("notes", cloudNotes.updatedAt);
      await saveNotes({
        skipCloud: true,
        saveTimestamp: false,
        silent: true,
      });
      renderNoteTabs();
      renderChecklistTabs();
      renderTextsView();
      renderNoteSearchResults();
      return cloudNotes.updatedAt;
    }
    if (domain === "templates") {
      const cloudTemplates = mergeLocalTemplatePaths(
        normalizeCloudTemplatesDoc(remoteDoc),
        templatesDb
      );
      templatesDb = {
        templates: cloudTemplates.templates,
        defaultTemplatesInstalled: cloudTemplates.defaultTemplatesInstalled,
        lastModified: cloudTemplates.updatedAt || cloudTemplates.lastModified,
      };
      touchLocalSyncTimestamp("templates", cloudTemplates.updatedAt);
      await saveTemplates({
        skipCloud: true,
        saveTimestamp: false,
        silent: true,
      });
      templatesDb = (await loadTemplates()) || templatesDb;
      renderTemplates();
      return cloudTemplates.updatedAt;
    }
    if (domain === "checklists") {
      const cloudChecklists = normalizeCloudChecklistsDoc(remoteDoc);
      checklistsDb = {
        checklists: cloudChecklists.checklists,
        templateOverrides: cloudChecklists.templateOverrides,
        lastModified: cloudChecklists.updatedAt || cloudChecklists.lastModified,
      };
      activeChecklistTabId =
        checklistsDb.checklists.find((checklist) => checklist.id === activeChecklistTabId)
          ?.id || checklistsDb.checklists[0]?.id || null;
      ensureActiveNotebookSelection(activeNotebookType);
      touchLocalSyncTimestamp("checklists", cloudChecklists.updatedAt);
      await saveChecklists({
        skipCloud: true,
        saveTimestamp: false,
        silent: true,
      });
      renderNoteTabs();
      renderChecklistTabs();
      renderTextsView();
      return cloudChecklists.updatedAt;
    }
  } finally {
    endCloudSyncApply();
  }
  return "";
}

async function pullUserState({ silent = false } = {}) {
  if (!cloudSyncState.enabled || !cloudSyncState.firebaseUid) return false;
  updateCloudSyncState({
    busy: true,
    status: "syncing",
    message: "Checking for remote updates...",
    error: "",
  });

  let backupCreated = false;
  const maybeBackup = async (reason, metadata = {}) => {
    if (backupCreated) return;
    backupCreated = true;
    await createCloudSyncBackup(reason, metadata);
  };

  try {
    const remote = await fetchRemoteUserState(cloudSyncState.firebaseUid);
    const domains = ["settings", "tasks", "notes", "templates", "checklists"];
    const syncedAtValues = [];

    for (const domain of domains) {
      const localDoc = getLocalCloudDoc(domain);
      const remoteDoc = remote[domain];
      const localHasData = hasMeaningfulCloudDoc(domain, localDoc);
      const remoteHasData = hasMeaningfulCloudDoc(domain, remoteDoc);
      const localUpdatedAt =
        normalizeIsoTimestamp(localDoc?.updatedAt) ||
        normalizeIsoTimestamp(localSyncTimestamps[domain]);
      const remoteUpdatedAt = normalizeIsoTimestamp(remoteDoc?.updatedAt);

      if (
        remoteHasData &&
        (!localHasData || isIsoAfter(remoteUpdatedAt, localUpdatedAt))
      ) {
        if (localHasData && isIsoAfter(remoteUpdatedAt, localUpdatedAt)) {
          await maybeBackup(`cloud-remote-${domain}`, {
            domain,
            localUpdatedAt,
            remoteUpdatedAt,
            firebaseUid: cloudSyncState.firebaseUid,
          });
        }
        const appliedAt = await applyRemoteCloudDoc(domain, remoteDoc);
        if (appliedAt) syncedAtValues.push(appliedAt);
        continue;
      }

      if (
        localHasData &&
        (!remoteHasData || isIsoAfter(localUpdatedAt, remoteUpdatedAt))
      ) {
        const pushedAt = await pushCloudDomain(domain);
        if (pushedAt) syncedAtValues.push(pushedAt);
      }
    }

    const localTimesheets = getLocalCloudDoc("timesheets");
    const remoteTimesheets = remote.timesheets || normalizeRemoteTimesheetsState({}, []);
    const localTimesheetsHasData = hasMeaningfulCloudDoc("timesheets", localTimesheets);
    const remoteTimesheetsHasData = hasMeaningfulCloudDoc("timesheets", remoteTimesheets);
    const localTimesheetsUpdatedAt =
      normalizeIsoTimestamp(localTimesheets?.updatedAt) ||
      normalizeIsoTimestamp(localSyncTimestamps.timesheets);
    const remoteTimesheetsUpdatedAt = normalizeIsoTimestamp(remoteTimesheets?.updatedAt);

    if (
      remoteTimesheetsHasData &&
      (!localTimesheetsHasData ||
        isIsoAfter(remoteTimesheetsUpdatedAt, localTimesheetsUpdatedAt))
    ) {
      if (
        localTimesheetsHasData &&
        isIsoAfter(remoteTimesheetsUpdatedAt, localTimesheetsUpdatedAt)
      ) {
        await maybeBackup("cloud-remote-timesheets", {
          domain: "timesheets",
          localUpdatedAt: localTimesheetsUpdatedAt,
          remoteUpdatedAt: remoteTimesheetsUpdatedAt,
          firebaseUid: cloudSyncState.firebaseUid,
        });
      }
      timesheetDb = {
        weeks: deepCloneJson(remoteTimesheets.weeks, {}),
        lastModified: remoteTimesheets.updatedAt || null,
      };
      cloudSyncRemoteTimesheetMeta = {
        updatedAt: remoteTimesheets.updatedAt || "",
        knownWeeks: remoteTimesheets.knownWeeks || [],
      };
      touchLocalSyncTimestamp("timesheets", remoteTimesheets.updatedAt);
      await saveTimesheets({
        skipCloud: true,
        saveTimestamp: false,
        silent: true,
      });
      renderTimesheets();
      await refreshTimesheetsInfo();
      if (remoteTimesheets.updatedAt) {
        syncedAtValues.push(remoteTimesheets.updatedAt);
      }
    } else if (
      localTimesheetsHasData &&
      (!remoteTimesheetsHasData ||
        isIsoAfter(localTimesheetsUpdatedAt, remoteTimesheetsUpdatedAt))
    ) {
      const pushedAt = await pushCloudTimesheetsState();
      if (pushedAt) syncedAtValues.push(pushedAt);
    } else {
      cloudSyncRemoteTimesheetMeta = {
        updatedAt: remoteTimesheets.updatedAt || "",
        knownWeeks: remoteTimesheets.knownWeeks || [],
      };
    }

    const lastSyncedAt = new Date().toISOString();
    await updateLocalCloudSyncMetadata(
      {
        enabled: true,
        firebaseUid: cloudSyncState.firebaseUid,
        migrationCompleted: true,
        lastSyncedAt,
      },
      { persist: true }
    );
    updateCloudSyncState({
      busy: false,
      enabled: true,
      status: "synced",
      message: lastSyncedAt
        ? `Last sync ${formatSyncTimestamp(lastSyncedAt)}`
        : "Cloud sync ready",
      error: "",
      lastSyncedAt,
    });
    return true;
  } catch (e) {
    if (!silent) {
      toast(e?.message || "Cloud sync failed.");
    }
    updateCloudSyncState({
      busy: false,
      status: "error",
      error: String(e?.message || "Cloud sync failed."),
    });
    throw e;
  }
}

async function bootstrapCloudSync({ session = null, silent = true } = {}) {
  if (cloudSyncInitPromise) return cloudSyncInitPromise;
  cloudSyncInitPromise = (async () => {
    await loadCloudSyncConfig();
    if (!cloudSyncState.configured) {
      updateCloudSyncState({
        enabled: false,
        status: "local-only",
        message: "Firebase is not configured.",
      });
      return { enabled: false };
    }
    if (!googleAuthState.signedIn) {
      updateCloudSyncState({
        enabled: false,
        signedIn: false,
        status: "idle",
        message: "Sign in to enable sync.",
      });
      return { enabled: false };
    }

    updateCloudSyncState({
      busy: true,
      signedIn: true,
      status: "syncing",
      message: "Connecting cloud sync...",
      error: "",
    });

    const user = await signInToCloud(session);
    if (!user?.uid) {
      updateCloudSyncState({
        busy: false,
        enabled: false,
        status: "idle",
        message: "Google is signed in locally.",
      });
      return { enabled: false };
    }

    await prepareLocalStateForUserSwitch(user.uid);
    updateCloudSyncState({
      enabled: true,
      signedIn: true,
      firebaseUid: user.uid,
    });
    await updateLocalCloudSyncMetadata(
      {
        enabled: true,
        firebaseUid: user.uid,
      },
      { persist: true }
    );
    await pullUserState({ silent });
    subscribeToUserState(user.uid);
    return { enabled: true, uid: user.uid };
  })()
    .catch((error) => {
      if (!silent) {
        toast(error?.message || "Cloud sync failed.");
      }
      updateCloudSyncState({
        busy: false,
        enabled: false,
        status: "error",
        error: String(error?.message || "Cloud sync failed."),
      });
      return { enabled: false, error };
    })
    .finally(() => {
      cloudSyncInitPromise = null;
    });
  return cloudSyncInitPromise;
}

async function signOutCloud({ preserveMetadata = true } = {}) {
  clearCloudSyncSubscriptions();
  Object.values(cloudSyncPushTimers).forEach((timer) => clearTimeout(timer));
  cloudSyncPushTimers = {};
  if (cloudSyncTimesheetsPushTimer) {
    clearTimeout(cloudSyncTimesheetsPushTimer);
    cloudSyncTimesheetsPushTimer = null;
  }
  if (firebaseAuthInstance) {
    try {
      await firebaseAuthInstance.signOut();
    } catch (e) {
      console.warn("Failed to sign out of Firebase:", e);
    }
  }
  cloudSyncRemoteTimesheetMeta = {
    updatedAt: "",
    knownWeeks: [],
  };
  await updateLocalCloudSyncMetadata(
    {
      enabled: false,
      migrationCompleted: false,
      ...(preserveMetadata ? {} : { firebaseUid: "", lastSyncedAt: "" }),
    },
    { persist: true }
  );
  updateCloudSyncState({
    enabled: false,
    busy: false,
    signedIn: false,
    status: "idle",
    message: "Sign in to enable sync.",
    error: "",
    ...(preserveMetadata
      ? {}
      : {
          firebaseUid: "",
          lastSyncedAt: "",
        }),
  });
}

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
  await persistUserSettingsLocally({ silent: true });
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

async function persistUserSettingsLocally({
  skipCloud = false,
  saveTimestamp = true,
  timestamp = new Date().toISOString(),
  silent = false,
} = {}) {
  const resolvedTimestamp =
    normalizeIsoTimestamp(timestamp) || new Date().toISOString();
  const comparableChanged =
    getCloudComparableFingerprint("settings") !==
    lastCloudComparableFingerprints.settings;
  if (saveTimestamp && comparableChanged) {
    touchLocalSyncTimestamp("settings", resolvedTimestamp);
  }
  try {
    const response = await window.pywebview.api.save_user_settings(userSettings);
    if (response?.status && response.status !== "success") {
      throw new Error(response.message || "Failed to save settings.");
    }
    syncCloudComparableFingerprint("settings");
    if (
      !skipCloud &&
      comparableChanged &&
      cloudSyncState.enabled &&
      !isCloudSyncApplying()
    ) {
      queueCloudStatePush("settings");
    }
    return true;
  } catch (e) {
    console.warn("Failed to save settings:", e);
    if (!silent) {
      toast("Could not save settings.");
    }
    return false;
  }
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

async function save({
  skipCloud = false,
  saveTimestamp = true,
  timestamp = new Date().toISOString(),
  silent = false,
} = {}) {
  syncPinnedProjectOrders(db, { seedMissing: true });
  db.forEach((project) => {
    syncProjectAttachmentFields(project);
    getProjectDeliverables(project).forEach((deliverable) => {
      syncDeliverableAttachmentFields(deliverable);
      syncDeliverableWorkItemFields(deliverable);
    });
  });
  const resolvedTimestamp =
    normalizeIsoTimestamp(timestamp) || new Date().toISOString();
  const comparableChanged =
    getCloudComparableFingerprint("tasks") !==
    lastCloudComparableFingerprints.tasks;
  if (saveTimestamp && comparableChanged) {
    touchLocalSyncTimestamp("tasks", resolvedTimestamp);
  }
  try {
    const response = await window.pywebview.api.save_tasks(db);
    if (response.status !== "success") throw new Error(response.message);
    syncCloudComparableFingerprint("tasks");
    if (
      !skipCloud &&
      comparableChanged &&
      cloudSyncState.enabled &&
      !isCloudSyncApplying()
    ) {
      queueCloudStatePush("tasks");
    }
    return true;
  } catch (e) {
    console.warn("Backend save failed:", e);
    if (silent) {
      return false;
    }
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

async function saveNotes({
  skipCloud = false,
  saveTimestamp = true,
  timestamp = new Date().toISOString(),
  silent = false,
} = {}) {
  const resolvedTimestamp =
    normalizeIsoTimestamp(timestamp) || new Date().toISOString();
  const comparableChanged =
    getCloudComparableFingerprint("notes") !==
    lastCloudComparableFingerprints.notes;
  if (saveTimestamp && comparableChanged) {
    touchLocalSyncTimestamp("notes", resolvedTimestamp);
  }
  try {
    const dataToSave = { tabs: noteTabs, keyed: {}, general: notesDb };
    const response = await window.pywebview.api.save_notes(dataToSave);
    if (response?.status && response.status !== "success") {
      throw new Error(response.message || "Failed to save notes.");
    }
    syncCloudComparableFingerprint("notes");
    if (
      !skipCloud &&
      comparableChanged &&
      cloudSyncState.enabled &&
      !isCloudSyncApplying()
    ) {
      queueCloudStatePush("notes");
    }
    return true;
  } catch (e) {
    console.warn("Backend notes save failed:", e);
    if (!silent) {
      toast("Failed to save notes.");
    }
    return false;
  }
}

async function loadUserSettings() {
  try {
    const storedSettings = await window.pywebview.api.get_user_settings();
    if (storedSettings) userSettings = { ...userSettings, ...storedSettings };
    userSettings.googleAuth =
      userSettings.googleAuth && typeof userSettings.googleAuth === "object"
        ? userSettings.googleAuth
        : null;
    delete userSettings.microsoftAuth;
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
    userSettings.enableUnderConstructionTools =
      userSettings.enableUnderConstructionTools === true;
    userSettings.separateDeliverableCompletionGroups =
      userSettings.separateDeliverableCompletionGroups !== false;
    userSettings.groupDeliverablesByProject =
      userSettings.groupDeliverablesByProject === true;
    userSettings.cloudSync = normalizeCloudSyncSettings(userSettings.cloudSync);
    syncProjectViewPreferencesFromSettings();
    updateCloudSyncState({
      enabled: userSettings.cloudSync.enabled,
      firebaseUid: userSettings.cloudSync.firebaseUid,
      lastSyncedAt: userSettings.cloudSync.lastSyncedAt,
    });
    syncCloudComparableFingerprint("settings");
    syncUnderConstructionToolsAvailability();
    renderGoogleAuthUi();
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
}

function areUnderConstructionToolsEnabled() {
  return userSettings.enableUnderConstructionTools === true;
}

function syncUnderConstructionToolsInputs() {
  setCheckboxValue(
    "settings_enableUnderConstructionTools",
    areUnderConstructionToolsEnabled()
  );
}

function syncUnderConstructionToolsAvailability() {
  const enabled = areUnderConstructionToolsEnabled();
  document
    .querySelectorAll('#tools-panel .tool-card[data-under-construction="true"]')
    .forEach((card) => {
      card.dataset.underConstructionEnabled = enabled ? "true" : "false";
      card.setAttribute("aria-disabled", enabled ? "false" : "true");
      card.title = enabled
        ? "Preview enabled. This tool is still under construction."
        : "Under construction. Enable preview access in Settings to use this tool.";
    });

  const warning = document.getElementById("toolsUnderConstructionWarning");
  if (warning) {
    warning.textContent = enabled
      ? "Preview access is enabled. These tools are still under construction and may be incomplete."
      : "These tools are currently being worked on and are not ready for use.";
  }
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
  setCheckboxValue(
    "settings_separateDeliverableCompletionGroups",
    userSettings.separateDeliverableCompletionGroups
  );
  setCheckboxValue(
    "settings_groupDeliverablesByProject",
    userSettings.groupDeliverablesByProject
  );

  syncCleanOptionsInputs();
  syncPublishOptionsInputs();
  syncFreezeOptionsInputs();
  syncThawOptionsInputs();
  syncWorkroomCadRoutingInputs();
  syncUnderConstructionToolsInputs();
  syncUnderConstructionToolsAvailability();
  renderGoogleAuthUi();

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
  const separateCompletionGroupsCheck = document.getElementById(
    "settings_separateDeliverableCompletionGroups"
  );
  if (separateCompletionGroupsCheck) {
    userSettings.separateDeliverableCompletionGroups =
      separateCompletionGroupsCheck.checked;
  }
  const groupByProjectCheck = document.getElementById(
    "settings_groupDeliverablesByProject"
  );
  if (groupByProjectCheck) {
    userSettings.groupDeliverablesByProject = groupByProjectCheck.checked;
  }
  syncProjectViewPreferencesFromSettings();
  await persistUserSettingsLocally();
}
const debouncedSaveUserSettings = debounce(saveUserSettings, 500);

function createId(prefix = "dlv") {
  if (window.crypto?.randomUUID) return `${prefix}_${crypto.randomUUID()}`;
  return `${prefix}_${Date.now().toString(36)}_${Math.random()
    .toString(36)
    .slice(2, 8)}`;
}

function normalizeTaskLinks(links) {
  return normalizeDeliverableLinks(links);
}

function normalizeTask(task) {
  const out =
    typeof task === "string"
      ? {
          text: task,
          done: false,
          pinned: false,
          attachments: [],
        }
      : {
          text: task?.text || "",
          done: !!task?.done,
          pinned: !!task?.pinned,
          attachments: normalizeAttachments(task?.attachments, {
            legacyLinks: task?.links,
            legacyEmailRefs: task?.emailRefs,
          }),
        };
  syncAttachmentOwnerCompatFields(out, {
    includeEmailRefs: true,
  });
  return {
    ...out,
    text: out.text || "",
    done: !!out.done,
    pinned: !!out.pinned,
  };
}

function normalizeDeliverableNoteItem(noteItem) {
  const out =
    typeof noteItem === "string"
      ? { text: noteItem, pinned: false, attachments: [] }
      : {
          text: noteItem?.text || "",
          pinned: !!noteItem?.pinned,
          attachments: normalizeAttachments(noteItem?.attachments, {
            legacyLinks: noteItem?.links,
            legacyEmailRefs: noteItem?.emailRefs,
          }),
        };
  syncAttachmentOwnerCompatFields(out, {
    includeEmailRefs: true,
  });
  return {
    ...out,
    text: out.text || "",
    pinned: !!out.pinned,
  };
}

function normalizeDeliverableNoteItems(noteItems = [], legacyNotes = "") {
  const normalized = [];
  const hasExplicitNoteItems = Array.isArray(noteItems);
  if (hasExplicitNoteItems) {
    noteItems.forEach((noteItem) => {
      const normalizedItem = normalizeDeliverableNoteItem(noteItem);
      if (!String(normalizedItem.text || "").trim()) return;
      normalized.push(normalizedItem);
    });
  }

  if (!hasExplicitNoteItems && String(legacyNotes || "").trim()) {
    normalized.push({
      text: String(legacyNotes || ""),
      pinned: false,
    });
  }

  return normalized;
}

function buildDeliverableNotesText(noteItems = [], fallback = "") {
  const normalized = Array.isArray(noteItems)
    ? noteItems
        .map(normalizeDeliverableNoteItem)
        .filter((noteItem) => String(noteItem.text || "").trim())
    : [];
  if (!normalized.length) return String(fallback || "");
  return normalized.map((noteItem) => noteItem.text).join("\n");
}

function syncDeliverableNoteFields(deliverable) {
  if (!deliverable || typeof deliverable !== "object") return deliverable;
  deliverable.noteItems = normalizeDeliverableNoteItems(
    deliverable.noteItems,
    deliverable.notes
  );
  deliverable.notes = buildDeliverableNotesText(
    deliverable.noteItems,
    deliverable.notes
  );
  return deliverable;
}

function syncDeliverableWorkItemFields(deliverable) {
  if (!deliverable || typeof deliverable !== "object") return deliverable;
  const nextTasks = [];
  (Array.isArray(deliverable.tasks) ? deliverable.tasks : []).forEach((task) => {
    const normalizedTask = normalizeTask(task);
    if (!String(normalizedTask.text || "").trim()) return;
    if (task && typeof task === "object" && !Array.isArray(task)) {
      task.text = normalizedTask.text;
      task.done = normalizedTask.done;
      task.pinned = normalizedTask.pinned;
      task.attachments = normalizedTask.attachments;
      task.links = normalizedTask.links;
      task.emailRefs = normalizedTask.emailRefs;
      task.emailRef = normalizedTask.emailRefs[0] || null;
      nextTasks.push(task);
      return;
    }
    nextTasks.push(normalizedTask);
  });
  deliverable.tasks = nextTasks;

  const nextNoteItems = [];
  (
    Array.isArray(deliverable.noteItems)
      ? deliverable.noteItems
      : normalizeDeliverableNoteItems([], deliverable.notes)
  ).forEach((noteItem) => {
    const normalizedNoteItem = normalizeDeliverableNoteItem(noteItem);
    if (!String(normalizedNoteItem.text || "").trim()) return;
    if (noteItem && typeof noteItem === "object" && !Array.isArray(noteItem)) {
      noteItem.text = normalizedNoteItem.text;
      noteItem.pinned = normalizedNoteItem.pinned;
      noteItem.attachments = normalizedNoteItem.attachments;
      noteItem.links = normalizedNoteItem.links;
      noteItem.emailRefs = normalizedNoteItem.emailRefs;
      noteItem.emailRef = normalizedNoteItem.emailRefs[0] || null;
      nextNoteItems.push(noteItem);
      return;
    }
    nextNoteItems.push(normalizedNoteItem);
  });
  deliverable.noteItems = nextNoteItems;
  return syncDeliverableNoteFields(deliverable);
}

function getPinnedFirstEntries(items = [], normalizeItem = (item) => item) {
  return (Array.isArray(items) ? items : [])
    .map((item, index) => ({ item: normalizeItem(item), index }))
    .filter(({ item }) => String(item?.text || "").trim())
    .sort((left, right) => {
      if (!!left.item?.pinned === !!right.item?.pinned) {
        return left.index - right.index;
      }
      return left.item?.pinned ? -1 : 1;
    });
}

function getPinnedDeliverableTaskEntries(deliverable) {
  return getPinnedFirstEntries(deliverable?.tasks, normalizeTask);
}

function getPinnedDeliverableNoteEntries(deliverable) {
  return getPinnedFirstEntries(
    deliverable?.noteItems,
    normalizeDeliverableNoteItem
  );
}

function getPinnedDeliverablePreviewItems(deliverable) {
  return [
    ...getPinnedDeliverableTaskEntries(deliverable)
      .filter(({ item }) => item.pinned)
      .map(({ item, index }) => ({
        type: "task",
        text: item.text,
        done: !!item.done,
        pinned: true,
        attachments: normalizeAttachments(item.attachments, {
          legacyLinks: item.links,
          legacyEmailRefs: item.emailRefs,
        }),
        links: normalizeTaskLinks(item.links),
        emailRefs: normalizeEmailRefs(item.emailRefs),
        index,
      })),
    ...getPinnedDeliverableNoteEntries(deliverable)
      .filter(({ item }) => item.pinned)
      .map(({ item, index }) => ({
        type: "note",
        text: item.text,
        pinned: true,
        attachments: normalizeAttachments(item.attachments, {
          legacyLinks: item.links,
          legacyEmailRefs: item.emailRefs,
        }),
        links: normalizeDeliverableLinks(item.links),
        emailRefs: normalizeEmailRefs(item.emailRefs),
        index,
      })),
  ];
}

function normalizePinnedProjectOrder(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return null;
  const normalized = Math.trunc(numeric);
  return normalized >= 0 ? normalized : null;
}

function compareProjectsByStableIdentity(a, b) {
  const idDiff = String(a?.id || "").localeCompare(String(b?.id || ""), undefined, {
    numeric: true,
    sensitivity: "base",
  });
  if (idDiff) return idDiff;

  const nameDiff = String(a?.name || "").localeCompare(
    String(b?.name || ""),
    undefined,
    {
      numeric: true,
      sensitivity: "base",
    }
  );
  if (nameDiff) return nameDiff;

  return String(a?.nick || "").localeCompare(String(b?.nick || ""), undefined, {
    numeric: true,
    sensitivity: "base",
  });
}

function buildLegacyPinnedProjectSortContextMap(items = []) {
  return new Map(
    (Array.isArray(items) ? items : []).map((project) => {
      const priority = getProjectListPriorityMeta(project);
      return [
        project,
        {
          activeAnchorDeliverable: priority.priorityDeliverable,
          anchorDueDate: priority.hasIncompleteActiveWork
            ? priority.sortDueDate
            : priority.fallbackDueDate,
          sortBucket: priority.sortBucket,
          fallbackDueDate: priority.fallbackDueDate,
        },
      ];
    })
  );
}

function compareProjectsByLegacyPinnedOrder(a, b, projectListContextMap = null) {
  const bucketDiff = compareProjectListSortBuckets(a, b, projectListContextMap);
  if (bucketDiff) return bucketDiff;
  const da = getProjectSortKey(a, projectListContextMap?.get(a));
  const dbb = getProjectSortKey(b, projectListContextMap?.get(b));
  if (!da && !dbb) return 0;
  if (!da) return 1;
  if (!dbb) return -1;
  const dueDiff = da - dbb;
  if (dueDiff) return dueDiff;
  return 0;
}

function getPinnedProjectsInManualOrder(items = db) {
  return (Array.isArray(items) ? items : [])
    .map((project, index) => ({
      project,
      index,
      pinnedOrder: normalizePinnedProjectOrder(project?.pinnedOrder),
    }))
    .filter(({ project }) => project?.pinned)
    .sort((left, right) => {
      if (left.pinnedOrder !== null && right.pinnedOrder !== null) {
        if (left.pinnedOrder !== right.pinnedOrder) {
          return left.pinnedOrder - right.pinnedOrder;
        }
      } else if (left.pinnedOrder !== null || right.pinnedOrder !== null) {
        return left.pinnedOrder !== null ? -1 : 1;
      }
      return left.index - right.index;
    })
    .map(({ project }) => project);
}

function syncPinnedProjectOrders(items = db, { seedMissing = false } = {}) {
  if (!Array.isArray(items)) return false;

  let changed = false;
  const orderedPinned = [];
  const missingPinned = [];

  items.forEach((project, index) => {
    if (!project || typeof project !== "object") return;

    const isPinned = project.pinned === true;
    if (project.pinned !== isPinned) {
      project.pinned = isPinned;
      changed = true;
    }

    const normalizedOrder = normalizePinnedProjectOrder(project.pinnedOrder);
    if (!isPinned) {
      if (project.pinnedOrder != null) changed = true;
      project.pinnedOrder = null;
      return;
    }

    if (project.pinnedOrder !== normalizedOrder) {
      project.pinnedOrder = normalizedOrder;
      changed = true;
    }

    const entry = { project, index };
    if (normalizedOrder === null) {
      missingPinned.push(entry);
      return;
    }
    orderedPinned.push({ ...entry, pinnedOrder: normalizedOrder });
  });

  orderedPinned.sort((left, right) => {
    if (left.pinnedOrder !== right.pinnedOrder) {
      return left.pinnedOrder - right.pinnedOrder;
    }
    return left.index - right.index;
  });

  if (missingPinned.length) {
    if (seedMissing) {
      const projectListContextMap = buildLegacyPinnedProjectSortContextMap(
        missingPinned.map(({ project }) => project)
      );
      missingPinned.sort((left, right) => {
        const legacyDiff = compareProjectsByLegacyPinnedOrder(
          left.project,
          right.project,
          projectListContextMap
        );
        if (legacyDiff) return legacyDiff;
        return left.index - right.index;
      });
    } else {
      missingPinned.sort((left, right) => left.index - right.index);
    }
  }

  orderedPinned
    .map(({ project }) => project)
    .concat(missingPinned.map(({ project }) => project))
    .forEach((project, index) => {
      if (project.pinnedOrder !== index) {
        project.pinnedOrder = index;
        changed = true;
      }
    });

  return changed;
}

function getNextPinnedProjectOrder(items = db) {
  return getPinnedProjectsInManualOrder(items).length;
}

function getLowestPinnedProjectOrder(projects = []) {
  let lowest = null;
  (Array.isArray(projects) ? projects : []).forEach((project) => {
    if (!project?.pinned) return;
    const pinnedOrder = normalizePinnedProjectOrder(project?.pinnedOrder);
    if (pinnedOrder === null) return;
    if (lowest === null || pinnedOrder < lowest) lowest = pinnedOrder;
  });
  return lowest;
}

function setProjectPinnedState(project, nextPinned, items = db) {
  if (!project || typeof project !== "object") return false;

  if (nextPinned) {
    project.pinned = true;
    if (normalizePinnedProjectOrder(project.pinnedOrder) === null) {
      project.pinnedOrder = getNextPinnedProjectOrder(items);
    }
  } else {
    project.pinned = false;
    project.pinnedOrder = null;
  }

  syncPinnedProjectOrders(items);
  return project.pinned === true;
}

function movePinnedProjectToTarget(project, targetProject, before = true, items = db) {
  if (!project?.pinned || !targetProject?.pinned || project === targetProject) {
    return false;
  }

  const orderedPinned = getPinnedProjectsInManualOrder(items);
  const fromIndex = orderedPinned.indexOf(project);
  const targetIndex = orderedPinned.indexOf(targetProject);
  if (fromIndex < 0 || targetIndex < 0 || fromIndex === targetIndex) return false;

  const [movedProject] = orderedPinned.splice(fromIndex, 1);
  let insertIndex = before ? targetIndex : targetIndex + 1;
  if (fromIndex < insertIndex) insertIndex -= 1;
  orderedPinned.splice(insertIndex, 0, movedProject);
  orderedPinned.forEach((pinnedProject, index) => {
    pinnedProject.pinned = true;
    pinnedProject.pinnedOrder = index;
  });
  syncPinnedProjectOrders(items);
  return true;
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

function normalizeDeliverableLinkEntry(entry) {
  if (!entry) return null;

  const rawInput =
    typeof entry === "string" ? entry : entry.raw || entry.url || entry.label || "";
  let raw = String(rawInput || "").trim();
  if (!raw) return null;

  const localPathFromUrl = fileUrlToLocalPath(raw);
  if (localPathFromUrl) {
    raw = normalizeWindowsPath(localPathFromUrl);
  } else if (/^[A-Za-z]:[\\/]/.test(raw) || /^\\\\/.test(raw)) {
    raw = normalizeWindowsPath(raw);
  }

  if (!raw) return null;

  const normalized = normalizeLink(raw);
  const label = String(
    typeof entry === "object" && entry?.label ? entry.label : normalized.label || raw
  ).trim();

  return {
    raw,
    url: normalized.url || normalized.raw || raw,
    label: label || normalized.label || raw,
  };
}

function hasMeaningfulDeliverableLinkLabel(link) {
  if (!link) return false;
  const label = String(link.label || "").trim();
  if (!label) return false;
  const raw = String(link.raw || "").trim();
  const url = String(link.url || "").trim();
  return label !== raw && label !== url;
}

function normalizeDeliverableLinks(links, legacyLinkPath = "") {
  const normalized = [];
  const seen = new Map();
  const candidates = Array.isArray(links) ? [...links] : [];

  if (legacyLinkPath) {
    candidates.push({ raw: legacyLinkPath });
  }

  candidates.forEach((entry) => {
    const next = normalizeDeliverableLinkEntry(entry);
    if (!next) return;

    const key = `${String(next.raw || "").toLowerCase()}|${String(next.url || "").toLowerCase()}`;
    if (seen.has(key)) {
      const existing = seen.get(key);
      if (
        hasMeaningfulDeliverableLinkLabel(next) &&
        !hasMeaningfulDeliverableLinkLabel(existing)
      ) {
        existing.label = next.label;
      } else if (!existing.label && next.label) {
        existing.label = next.label;
      }
      return;
    }

    seen.set(key, next);
    normalized.push(next);
  });

  return normalized;
}

function normalizeDeliverable(deliverable = {}) {
  const attachments = normalizeAttachments(deliverable.attachments, {
    legacyLinks: normalizeDeliverableLinks(
      deliverable.links,
      deliverable.linkPath || ""
    ),
    legacyEmailRefs: deliverable.emailRefs,
    legacyEmailRef: deliverable.emailRef,
  });
  const emailRefs = buildLegacyEmailRefsFromAttachments(attachments);
  const out = {
    id: deliverable.id || createId("dlv"),
    name: String(deliverable.name || "").trim(),
    due: String(deliverable.due || "").trim(),
    notes: String(deliverable.notes || ""),
    noteItems: normalizeDeliverableNoteItems(
      deliverable.noteItems,
      deliverable.notes || ""
    ),
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
    attachments,
    emailRefs,
    emailRef: emailRefs[0] || null,
    links: buildLegacyLinksFromAttachments(attachments),
    active: deliverable.active === true,
    workroomCadDiscipline: normalizeWorkroomCadDiscipline(
      deliverable.workroomCadDiscipline,
      ""
    ),
    workroomPhase: normalizeWorkroomPhase(deliverable.workroomPhase),
    workroomReturnType: normalizeWorkroomReturnType(deliverable.workroomReturnType),
    appliedChecklists: Array.isArray(deliverable.appliedChecklists)
      ? deliverable.appliedChecklists
          .map((instance) => normalizeChecklistInstance(instance))
          .filter(Boolean)
      : [],
    checklistContext: normalizeChecklistContext(
      deliverable.checklistContext || createDefaultChecklistContext()
    ),
  };
  migrateStatusFields(out);
  syncStatusArrays(out);
  syncDeliverableAttachmentFields(out);
  syncDeliverableWorkItemFields(out);
  if (isFinished(out)) {
    out.tasks.forEach((task) => {
      task.done = true;
    });
  }
  return out;
}

function createDeliverable(seed = {}) {
  const seedAttachments = normalizeAttachments(seed.attachments, {
    legacyLinks: normalizeDeliverableLinks(seed.links, seed.linkPath || ""),
    legacyEmailRefs: seed.emailRefs,
    legacyEmailRef: seed.emailRef,
  });
  return normalizeDeliverable({
    id: seed.id || createId("dlv"),
    name: seed.name || "",
    due: seed.due || "",
    notes: seed.notes || "",
    noteItems: seed.noteItems || [],
    tasks: seed.tasks || [],
    statuses: seed.statuses || [],
    statusTags: seed.statusTags || [],
    status: seed.status || "",
    attachments: seedAttachments,
    active: seed.active !== false,
    workroomCadDiscipline: seed.workroomCadDiscipline || "",
    workroomPhase: seed.workroomPhase || "pre_design",
    workroomReturnType: seed.workroomReturnType || "",
    appliedChecklists: seed.appliedChecklists || [],
    checklistContext: seed.checklistContext || createDefaultChecklistContext(),
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
  const path = normalizeProjectPath(project?.path || "").toLowerCase();
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
  if (!base.localProjectPath && incoming.localProjectPath) {
    base.localProjectPath = incoming.localProjectPath;
  }
  if (!base.workroomRootPath && incoming.workroomRootPath) {
    base.workroomRootPath = incoming.workroomRootPath;
  }
  if (!base.notes && incoming.notes) base.notes = incoming.notes;
  base.refs = mergeRefs(base.refs || [], incoming.refs || []);
  base.attachments = normalizeAttachments(
    [
      ...(Array.isArray(base.attachments) ? base.attachments : []),
      ...(Array.isArray(incoming.attachments) ? incoming.attachments : []),
    ],
    {
      legacyLinks: [
        ...(Array.isArray(base.links) ? base.links : []),
        ...(Array.isArray(incoming.links) ? incoming.links : []),
      ],
    }
  );
  syncProjectAttachmentFields(base);
  if (!base.lightingSchedule && incoming.lightingSchedule)
    base.lightingSchedule = incoming.lightingSchedule;
  if (!base.title24 && incoming.title24) base.title24 = incoming.title24;
  const baseChecklistProfiles = normalizeProjectChecklistProfiles(base.checklistProfiles);
  const incomingChecklistProfiles = normalizeProjectChecklistProfiles(incoming.checklistProfiles);
  const baseScopeProfile =
    baseChecklistProfiles[ACIES_ELECTRICAL_SCOPE_PROFILE_KEY] ||
    createDefaultAciesElectricalChecklistProfile();
  const incomingScopeProfile =
    incomingChecklistProfiles[ACIES_ELECTRICAL_SCOPE_PROFILE_KEY] ||
    createDefaultAciesElectricalChecklistProfile();
  const mergedScopeAnswers = createEmptyAciesElectricalChecklistAnswers();
  ACIES_ELECTRICAL_SCOPE_FIELD_KEYS.forEach((fieldKey) => {
    mergedScopeAnswers[fieldKey] =
      baseScopeProfile.answers?.[fieldKey] ||
      incomingScopeProfile.answers?.[fieldKey] ||
      "";
  });
  base.checklistProfiles = {
    ...incomingChecklistProfiles,
    ...baseChecklistProfiles,
    [ACIES_ELECTRICAL_SCOPE_PROFILE_KEY]: {
      schemaVersion: CHECKLIST_PROFILE_SCHEMA_VERSION,
      answers: mergedScopeAnswers,
      updatedAtUtc:
        baseScopeProfile.updatedAtUtc || incomingScopeProfile.updatedAtUtc || "",
    },
  };
  if (!Array.isArray(base.deliverables)) base.deliverables = [];
  if (Array.isArray(incoming.deliverables)) {
    base.deliverables.push(
      ...incoming.deliverables.map((deliverable) =>
        normalizeDeliverable({
          ...deliverable,
          active: true,
        })
      )
    );
  }
  if (!base.overviewDeliverableId && incoming.overviewDeliverableId)
    base.overviewDeliverableId = incoming.overviewDeliverableId;
  base.pinned = base?.pinned === true || incoming?.pinned === true;
  base.pinnedOrder = base.pinned
    ? getLowestPinnedProjectOrder([base, incoming])
    : null;
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
    path: normalizeProjectPath(legacy?.path || ""),
    localProjectPath: "",
    workroomRootPath: "",
    notes: "",
    refs: Array.isArray(legacy?.refs)
      ? legacy.refs.map(normalizeRef).filter(Boolean)
      : [],
    attachments: normalizeAttachments(legacy?.attachments, {
      legacyLinks: legacy?.links,
    }),
    links: [],
    deliverables: [deliverable],
    overviewDeliverableId: deliverable.id,
    pinned: !!legacy?.pinned,
    pinnedOrder: normalizePinnedProjectOrder(legacy?.pinnedOrder),
    lightingSchedule: legacy?.lightingSchedule || null,
    title24: legacy?.title24 || null,
  };
}

function normalizeProject(project) {
  if (!project) return null;
  if (!Array.isArray(project.deliverables) && isLegacyProject(project)) {
    return normalizeProject(convertLegacyProject(project));
  }
  const attachments = normalizeAttachments(project.attachments, {
    legacyLinks: project.links,
  });
  const out = {
    ...project,
    id: String(project.id || "").trim(),
    name: String(project.name || "").trim(),
    nick: String(project.nick || "").trim(),
    path: normalizeProjectPath(project.path || ""),
    localProjectPath: normalizeWindowsPath(project.localProjectPath || ""),
    workroomRootPath: normalizeWindowsPath(project.workroomRootPath || ""),
    notes: project.notes || "",
    refs: Array.isArray(project.refs)
      ? project.refs.map(normalizeRef).filter(Boolean)
      : [],
    attachments,
    links: buildLegacyLinksFromAttachments(attachments),
    pinned: project?.pinned === true,
    pinnedOrder: normalizePinnedProjectOrder(project?.pinnedOrder),
    deliverables: Array.isArray(project.deliverables)
      ? project.deliverables.map(normalizeDeliverable)
      : [],
    checklistProfiles: normalizeProjectChecklistProfiles(
      project.checklistProfiles || { [ACIES_ELECTRICAL_SCOPE_PROFILE_KEY]: project.checklistProfile }
    ),
    lightingSchedule: normalizeLightingSchedule(
      project.lightingSchedule || createDefaultLightingSchedule()
    ),
    title24: normalizeTitle24(project.title24 || createDefaultTitle24()),
  };
  syncProjectAttachmentFields(out);
  if (!out.pinned) out.pinnedOrder = null;
  if (!out.deliverables.length) out.deliverables = [createDeliverable()];
  syncProjectActiveDeliverables(out, {
    fallbackActiveId: String(project.overviewDeliverableId || "").trim(),
  });
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

function isDeliverableActive(deliverable) {
  return deliverable?.active === true;
}

function getProjectActiveDeliverables(project) {
  return getProjectDeliverables(project).filter((deliverable) =>
    isDeliverableActive(deliverable)
  );
}

function getActiveAnchorDeliverable(project) {
  const deliverables = getProjectDeliverables(project);
  if (!deliverables.length) return null;
  const activeDeliverables = getProjectActiveDeliverables(project);
  const activeWithDue = activeDeliverables.filter((deliverable) =>
    parseDueStr(deliverable?.due)
  );
  if (activeWithDue.length) {
    return activeWithDue.sort(compareDeliverablesByDue)[0];
  }
  if (activeDeliverables.length) return activeDeliverables[0];
  return deliverables[0];
}

function syncProjectActiveDeliverables(project, { fallbackActiveId = "" } = {}) {
  if (!project) return project;
  const deliverables = getProjectDeliverables(project);
  if (!deliverables.length) {
    const deliverable = createDeliverable();
    project.deliverables = [deliverable];
    project.overviewDeliverableId = deliverable.id;
    return project;
  }

  const normalizedFallbackId = String(fallbackActiveId || "").trim();
  const hasExplicitActiveDeliverables = deliverables.some((deliverable) =>
    isDeliverableActive(deliverable)
  );
  if (!hasExplicitActiveDeliverables && normalizedFallbackId) {
    deliverables.forEach((deliverable) => {
      deliverable.active = deliverable.id === normalizedFallbackId;
    });
  }
  if (!deliverables.some((deliverable) => deliverable?.active)) {
    deliverables[0].active = true;
  }

  const activeAnchorDeliverable = getActiveAnchorDeliverable(project);
  project.overviewDeliverableId = activeAnchorDeliverable?.id || deliverables[0]?.id || "";
  return project;
}

function projectNeedsActiveMigration(sourceProject, normalizedProject) {
  if (
    String(sourceProject?.overviewDeliverableId || "").trim() !==
    String(normalizedProject?.overviewDeliverableId || "").trim()
  ) {
    return true;
  }
  const sourceDeliverables = Array.isArray(sourceProject?.deliverables)
    ? sourceProject.deliverables
    : [];
  const normalizedDeliverables = Array.isArray(normalizedProject?.deliverables)
    ? normalizedProject.deliverables
    : [];
  if (sourceDeliverables.length !== normalizedDeliverables.length) return true;
  return normalizedDeliverables.some(
    (deliverable, index) =>
      (sourceDeliverables[index]?.active === true) !== (deliverable?.active === true)
  );
}

function autoSetPrimary(project) {
  if (!project || !userSettings.autoPrimary) return;
  const latestId = getLatestDueDeliverableId(project.deliverables);
  if (!latestId) return;
  const latestDeliverable = getProjectDeliverables(project).find(
    (deliverable) => deliverable.id === latestId
  );
  if (latestDeliverable) latestDeliverable.active = true;
  syncProjectActiveDeliverables(project, { fallbackActiveId: latestId });
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
      if (project?.path !== normalizeWindowsPath(item?.path || "")) {
        changed = true;
      }
      if (projectNeedsActiveMigration(item, project)) {
        changed = true;
      }
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
        const beforeOverviewId = normalized.overviewDeliverableId;
        syncProjectActiveDeliverables(normalized, {
          fallbackActiveId: beforeOverviewId,
        });
        if (normalized.overviewDeliverableId !== beforeOverviewId) {
          changed = true;
        }
      }
      return normalized;
    })
    .filter(Boolean);
  if (syncPinnedProjectOrders(merged, { seedMissing: true })) {
    changed = true;
  }
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
      setProjectPinnedState(project, true, db);
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

function resetCopyProjectLocallyDialogState() {
  copyProjectLocallyDialogState = {
    ...DEFAULT_COPY_PROJECT_LOCALLY_DIALOG_STATE,
    sync: createDefaultLocalProjectManagerSyncState(),
    folderOptions: [],
  };
}

function formatCopyProjectLocallySizeLabel(sizeBytes) {
  const parsed = Number(sizeBytes);
  if (!Number.isFinite(parsed) || parsed < 0) return "Size unavailable";
  const gigabyte = 1024 ** 3;
  const megabyte = 1024 ** 2;
  return parsed >= gigabyte
    ? `${(parsed / gigabyte).toFixed(1)} GB`
    : `${(parsed / megabyte).toFixed(1)} MB`;
}

function normalizeCopyProjectLocallyFolderOption(option, index = 0) {
  const name = String(option?.name || "").trim();
  const path = String(option?.path || "").trim();
  const sizeBytesValue = Number(option?.sizeBytes);
  const sizeStatus =
    String(option?.sizeStatus || "")
      .trim()
      .toLowerCase() === "available"
      ? "available"
      : "unavailable";
  const fullSizeBytes =
    Number.isFinite(sizeBytesValue) && sizeStatus === "available"
      ? sizeBytesValue
      : null;
  const fullSizeLabel =
    String(option?.sizeLabel || "").trim() ||
    (sizeStatus === "available"
      ? formatCopyProjectLocallySizeLabel(sizeBytesValue)
      : "Size unavailable");
  const selectedByDefault = option?.selectedByDefault === true;
  return {
    id: `${name || "folder"}::${index}`,
    name,
    path,
    fullSizeBytes,
    fullSizeLabel,
    fullSizeStatus: sizeStatus,
    selectedByDefault,
    selectionMode: selectedByDefault ? "all" : "none",
    hasChildFolders: option?.hasChildFolders === true,
    childrenLoaded: option?.childrenLoaded === true,
    childrenExpanded: false,
    childrenLoading: false,
    childFolderOptions: [],
    parentDirectFilesSizeBytes: null,
    parentDirectFilesSizeLabel: "Size unavailable",
    parentDirectFilesSizeStatus: "unavailable",
  };
}

function normalizeCopyProjectLocallyChildFolderOption(
  parentOption,
  option,
  index = 0
) {
  const name = String(option?.name || "").trim();
  const path = String(option?.path || "").trim();
  const relativePath = String(option?.relativePath || "").trim();
  const sizeBytesValue = Number(option?.sizeBytes);
  const sizeStatus =
    String(option?.sizeStatus || "")
      .trim()
      .toLowerCase() === "available"
      ? "available"
      : "unavailable";
  return {
    id: `${parentOption?.id || "folder"}::child::${name || index}`,
    name,
    path,
    relativePath,
    sizeBytes:
      Number.isFinite(sizeBytesValue) && sizeStatus === "available"
        ? sizeBytesValue
        : null,
    sizeLabel:
      String(option?.sizeLabel || "").trim() ||
      (sizeStatus === "available"
        ? formatCopyProjectLocallySizeLabel(sizeBytesValue)
        : "Size unavailable"),
    sizeStatus,
    selected: parentOption?.selectionMode === "all",
  };
}

function getCopyProjectLocallyFolderOption(folderId) {
  return copyProjectLocallyDialogState.folderOptions.find(
    (option) => option.id === folderId
  );
}

function getCopyProjectLocallyChildFolderOption(folderId, childId) {
  const parentOption = getCopyProjectLocallyFolderOption(folderId);
  if (!parentOption) return { parentOption: null, childOption: null };
  const childOption = parentOption.childFolderOptions.find(
    (option) => option.id === childId
  );
  return { parentOption, childOption };
}

function getCopyProjectLocallyFolderDisplaySize(option) {
  if (!option || option.selectionMode !== "subset" || !option.childrenLoaded) {
    return {
      sizeBytes: option?.fullSizeBytes ?? null,
      sizeLabel: option?.fullSizeLabel || "Size unavailable",
      sizeStatus: option?.fullSizeStatus || "unavailable",
    };
  }

  let totalBytes = 0;
  let hasKnownSize = false;
  let hasUnavailableSize = false;

  if (option.parentDirectFilesSizeStatus === "available") {
    totalBytes += Number(option.parentDirectFilesSizeBytes) || 0;
    hasKnownSize = true;
  } else {
    hasUnavailableSize = true;
  }

  option.childFolderOptions.forEach((childOption) => {
    if (!childOption.selected) return;
    if (childOption.sizeStatus === "available") {
      totalBytes += Number(childOption.sizeBytes) || 0;
      hasKnownSize = true;
      return;
    }
    hasUnavailableSize = true;
  });

  if (!hasKnownSize && hasUnavailableSize) {
    return {
      sizeBytes: null,
      sizeLabel: "Size unavailable",
      sizeStatus: "unavailable",
    };
  }

  const formattedLabel = formatCopyProjectLocallySizeLabel(totalBytes);
  return {
    sizeBytes: totalBytes,
    sizeLabel: hasUnavailableSize ? `${formattedLabel} (partial)` : formattedLabel,
    sizeStatus: hasUnavailableSize ? "partial" : "available",
  };
}

function buildCopyProjectLocallySelectionPayload() {
  const selectedOptions = copyProjectLocallyDialogState.folderOptions.filter(
    (option) => option.selectionMode !== "none"
  );

  if (!selectedOptions.length) {
    return {
      hasSelection: false,
      selectedFolderNames: [],
      selectedFolderRequests: null,
    };
  }

  const hasSubsetSelections = selectedOptions.some(
    (option) => option.selectionMode === "subset"
  );
  if (!hasSubsetSelections) {
    return {
      hasSelection: true,
      selectedFolderNames: selectedOptions.map((option) => option.name).filter(Boolean),
      selectedFolderRequests: null,
    };
  }

  return {
    hasSelection: true,
    selectedFolderNames: null,
    selectedFolderRequests: selectedOptions
      .map((option) => {
        if (!option?.name) return null;
        if (option.selectionMode === "subset") {
          return {
            name: option.name,
            mode: "subset",
            selectedChildNames: option.childFolderOptions
              .filter((childOption) => childOption.selected)
              .map((childOption) => childOption.name)
              .filter(Boolean),
            includeParentRootFiles: true,
          };
        }
        return {
          name: option.name,
          mode: "all",
        };
      })
      .filter(Boolean),
  };
}

function syncCopyProjectLocallyRowSelectionState(row, isSelected, selectionMode = "") {
  if (!row) return;
  const selected = isSelected === true;
  row.dataset.selected = selected ? "true" : "false";
  row.dataset.selectionMode = String(selectionMode || "").trim().toLowerCase();
  row.classList.toggle("is-selected", selected);
  row.classList.toggle("is-partial", selectionMode === "subset");
}

function syncCopyProjectLocallyParentCheckboxState(input, option) {
  if (!input || !option) return;
  input.checked = option.selectionMode === "all";
  input.indeterminate = option.selectionMode === "subset";
  input.setAttribute(
    "aria-checked",
    option.selectionMode === "subset"
      ? "mixed"
      : option.selectionMode === "all"
        ? "true"
        : "false"
  );
}

function createCopyProjectLocallyFolderExpandButton(option) {
  if (option.hasChildFolders !== true) {
    return el("span", {
      className: "copy-project-locally-expand-spacer",
      "aria-hidden": "true",
    });
  }

  const isLoading = option.childrenLoading === true;
  const isExpanded = option.childrenExpanded === true;
  const label = isLoading
    ? `Loading subfolders in ${option.name || "folder"}`
    : isExpanded
      ? `Hide subfolders in ${option.name || "folder"}`
      : `Show subfolders in ${option.name || "folder"}`;
  return el("button", {
    type: "button",
    className: "copy-project-locally-expand-btn",
    disabled: isLoading,
    textContent: isLoading ? "..." : isExpanded ? "-" : "+",
    title: label,
    "aria-label": label,
    "aria-expanded": String(isExpanded),
    "aria-controls": `copyProjectLocallyChildren-${option.id}`,
    "data-copy-project-expand-btn": "true",
    "data-folder-id": option.id,
  });
}

function createCopyProjectLocallyParentRow(option) {
  const sizeInfo = getCopyProjectLocallyFolderDisplaySize(option);
  const checkboxInput = el("input", {
    className: "copy-project-locally-checkbox-input",
    type: "checkbox",
    checked: option.selectionMode === "all",
    "data-copy-project-parent-checkbox": "true",
    "data-folder-id": option.id,
    "aria-label": `Copy ${option.name || "folder"}`,
  });
  syncCopyProjectLocallyParentCheckboxState(checkboxInput, option);

  const row = el(
    "div",
    {
      className: "copy-project-locally-row copy-project-locally-parent-row",
      role: "listitem",
    },
    [
      createCopyProjectLocallyFolderExpandButton(option),
      el("label", { className: "copy-project-locally-row-main" }, [
        el("span", { className: "custom-check copy-project-locally-checkbox" }, [
          checkboxInput,
          el("span", {
            className: "checkmark",
            "aria-hidden": "true",
          }),
        ]),
        el("div", { className: "copy-project-locally-row-content" }, [
          el("div", {
            className: "copy-project-locally-row-title",
            textContent: option.name || "Untitled Folder",
          }),
        ]),
        el("div", {
          className: "copy-project-locally-row-size",
          textContent: sizeInfo.sizeLabel || "Size unavailable",
        }),
      ]),
    ]
  );

  syncCopyProjectLocallyRowSelectionState(
    row,
    option.selectionMode !== "none",
    option.selectionMode
  );
  return row;
}

function createCopyProjectLocallyChildRow(parentOption, childOption) {
  const row = el(
    "label",
    {
      className: "copy-project-locally-row copy-project-locally-child-row",
      role: "listitem",
    },
    [
      el("span", { className: "custom-check copy-project-locally-checkbox" }, [
        el("input", {
          className: "copy-project-locally-checkbox-input",
          type: "checkbox",
          checked: childOption.selected === true,
          "data-copy-project-child-checkbox": "true",
          "data-folder-id": parentOption.id,
          "data-child-id": childOption.id,
          "aria-label": `Copy ${childOption.relativePath || childOption.name || "subfolder"}`,
        }),
        el("span", {
          className: "checkmark",
          "aria-hidden": "true",
        }),
      ]),
      el("div", { className: "copy-project-locally-row-content" }, [
        el("div", {
          className: "copy-project-locally-row-title",
          textContent: childOption.name || "Untitled Folder",
        }),
      ]),
      el("div", {
        className: "copy-project-locally-row-size",
        textContent: childOption.sizeLabel || "Size unavailable",
      }),
    ]
  );

  syncCopyProjectLocallyRowSelectionState(row, childOption.selected === true, "child");
  return row;
}

function createCopyProjectLocallyFolderListItem(option) {
  const group = el("div", {
    className: "copy-project-locally-group",
    "data-folder-id": option.id,
  });
  group.appendChild(createCopyProjectLocallyParentRow(option));

  if (option.hasChildFolders !== true) {
    return group;
  }

  const childList = el("div", {
    id: `copyProjectLocallyChildren-${option.id}`,
    className: "copy-project-locally-child-list",
    hidden: option.childrenExpanded !== true,
  });

  if (option.childrenLoading === true) {
    childList.appendChild(
      el("div", {
        className: "copy-project-locally-child-status muted tiny",
        textContent: "Loading subfolders...",
      })
    );
  } else if (option.childrenLoaded === true && option.childFolderOptions.length) {
    option.childFolderOptions.forEach((childOption) =>
      childList.appendChild(createCopyProjectLocallyChildRow(option, childOption))
    );
  } else if (option.childrenLoaded === true) {
    childList.appendChild(
      el("div", {
        className: "copy-project-locally-child-status muted tiny",
        textContent: "No subfolders found.",
      })
    );
  }

  group.appendChild(childList);
  return group;
}

function renderCopyProjectLocallyFolderList(container, items) {
  if (!container) return;
  container.replaceChildren();
  if (!items.length) {
    container.appendChild(
      el("div", {
        className: "deliverable-notepad-empty",
        textContent: "No subfolders were found in the selected server project folder.",
      })
    );
    return;
  }
  items.forEach((item) => container.appendChild(createCopyProjectLocallyFolderListItem(item)));
}

function updateCopyProjectLocallyDialogSummary() {
  const summaryEl = document.getElementById("copyProjectLocallySummary");
  const confirmBtn = document.getElementById("copyProjectLocallyConfirmBtn");
  const selectionPayload = buildCopyProjectLocallySelectionPayload();
  const copyReady =
    copyProjectLocallyDialogState.copyPreviewLoaded === true &&
    copyProjectLocallyDialogState.localProjectExists !== true &&
    copyProjectLocallyDialogState.folderOptions.length > 0;
  if (confirmBtn) confirmBtn.disabled = !copyReady || !selectionPayload.hasSelection;

  if (!summaryEl) return;
  if (copyProjectLocallyDialogState.copyPreviewLoaded !== true) {
    summaryEl.textContent = "Choose a server project folder to start.";
    return;
  }
  if (copyProjectLocallyDialogState.localProjectExists === true) {
    summaryEl.textContent = "Local copy already exists.";
    return;
  }
  if (!copyProjectLocallyDialogState.folderOptions.length) {
    summaryEl.textContent = "No subfolders available to copy.";
    return;
  }
  if (!selectionPayload.hasSelection) {
    summaryEl.textContent = "No folders selected.";
    return;
  }

  const selectedOptions = copyProjectLocallyDialogState.folderOptions.filter(
    (option) => option.selectionMode !== "none"
  );
  let totalBytes = 0;
  let partialCount = 0;
  selectedOptions.forEach((option) => {
    const sizeInfo = getCopyProjectLocallyFolderDisplaySize(option);
    if (sizeInfo.sizeStatus === "partial" || sizeInfo.sizeStatus === "unavailable") {
      partialCount += 1;
    }
    if (Number.isFinite(sizeInfo.sizeBytes) && Number(sizeInfo.sizeBytes) >= 0) {
      totalBytes += Number(sizeInfo.sizeBytes);
    }
  });

  const folderLabel = selectedOptions.length === 1 ? "folder" : "folders";
  summaryEl.textContent = `${selectedOptions.length} ${folderLabel} selected - ${formatCopyProjectLocallySizeLabel(
    totalBytes
  )}${partialCount ? " (partial)" : ""}`;
}

function setCopyProjectLocallyFolderSelectionMode(option, mode) {
  if (!option) return;
  option.selectionMode = mode === "all" ? "all" : mode === "subset" ? "subset" : "none";
  if (!option.childrenLoaded) return;
  if (option.selectionMode === "all") {
    option.childFolderOptions.forEach((childOption) => {
      childOption.selected = true;
    });
    return;
  }
  if (option.selectionMode === "none") {
    option.childFolderOptions.forEach((childOption) => {
      childOption.selected = false;
    });
  }
}

function syncCopyProjectLocallyParentModeFromChildren(option) {
  if (!option || !option.childrenLoaded) return;
  const selectedCount = option.childFolderOptions.filter(
    (childOption) => childOption.selected
  ).length;
  if (selectedCount === option.childFolderOptions.length && option.childFolderOptions.length) {
    option.selectionMode = "all";
    return;
  }
  option.selectionMode = "subset";
}

function handleCopyProjectLocallyParentCheckboxChange(folderId, isChecked) {
  const option = getCopyProjectLocallyFolderOption(folderId);
  if (!option) return;
  setCopyProjectLocallyFolderSelectionMode(option, isChecked ? "all" : "none");
  renderCopyProjectLocallyDialog();
}

function handleCopyProjectLocallyChildCheckboxChange(folderId, childId, isChecked) {
  const { parentOption, childOption } = getCopyProjectLocallyChildFolderOption(
    folderId,
    childId
  );
  if (!parentOption || !childOption) return;
  childOption.selected = isChecked === true;
  syncCopyProjectLocallyParentModeFromChildren(parentOption);
  renderCopyProjectLocallyDialog();
}

function getCopyProjectLocallySourceArgs() {
  return {
    serverProjectPath: copyProjectLocallyDialogState.resolvedFromWorkroom
      ? null
      : copyProjectLocallyDialogState.serverProjectPath || "",
    launchContext: copyProjectLocallyDialogState.launchContext || null,
  };
}

async function toggleCopyProjectLocallyFolderChildren(folderId) {
  const option = getCopyProjectLocallyFolderOption(folderId);
  if (!option || option.hasChildFolders !== true || option.childrenLoading === true) {
    return;
  }

  if (option.childrenLoaded) {
    option.childrenExpanded = option.childrenExpanded !== true;
    renderCopyProjectLocallyDialog();
    return;
  }

  option.childrenLoading = true;
  option.childrenExpanded = true;
  renderCopyProjectLocallyDialog();

  try {
    const sourceArgs = getCopyProjectLocallySourceArgs();
    const result = await window.pywebview.api.preview_copy_project_locally_child_folders(
      sourceArgs.serverProjectPath,
      option.name,
      sourceArgs.launchContext
    );

    if (result?.status !== "success") {
      throw new Error(result?.message || "Failed to load project subfolders.");
    }

    option.parentDirectFilesSizeBytes =
      Number.isFinite(Number(result?.parentDirectFilesSizeBytes))
        ? Number(result.parentDirectFilesSizeBytes)
        : null;
    option.parentDirectFilesSizeLabel =
      String(result?.parentDirectFilesSizeLabel || "").trim() || "Size unavailable";
    option.parentDirectFilesSizeStatus =
      String(result?.parentDirectFilesSizeStatus || "").trim().toLowerCase() === "available"
        ? "available"
        : "unavailable";
    option.childFolderOptions = Array.isArray(result?.childFolderOptions)
      ? result.childFolderOptions.map((childOption, index) =>
          normalizeCopyProjectLocallyChildFolderOption(option, childOption, index)
        )
      : [];
    option.childrenLoaded = true;
  } catch (error) {
    option.childrenExpanded = false;
    toast(error?.message || "Failed to load project subfolders.", 7000);
  } finally {
    option.childrenLoading = false;
    renderCopyProjectLocallyDialog();
  }
}

function applyCopyProjectLocallySelectionMode(mode) {
  copyProjectLocallyDialogState.folderOptions.forEach((option) => {
    if (mode === "all") {
      setCopyProjectLocallyFolderSelectionMode(option, "all");
      return;
    }
    if (mode === "none") {
      setCopyProjectLocallyFolderSelectionMode(option, "none");
      return;
    }
    setCopyProjectLocallyFolderSelectionMode(
      option,
      option.selectedByDefault ? "all" : "none"
    );
  });
  renderCopyProjectLocallyDialog();
}

function renderCopyProjectLocallyDialog() {
  const activeTab =
    copyProjectLocallyDialogState.activeTab === "sync" ? "sync" : "copy";
  const copyTabBtn = document.getElementById("localProjectManagerCopyTabBtn");
  const syncTabBtn = document.getElementById("localProjectManagerSyncTabBtn");
  const copyPanel = document.getElementById("copyProjectLocallyCopyPanel");
  const syncPanel = document.getElementById("copyProjectLocallySyncPanel");
  const projectNameEl = document.getElementById("copyProjectLocallyProjectName");
  const serverPathEl = document.getElementById("copyProjectLocallyServerPath");
  const localPathEl = document.getElementById("copyProjectLocallyLocalPath");
  const listEl = document.getElementById("copyProjectLocallyFolderList");
  const chooseFolderBtn = document.getElementById("copyProjectLocallyChooseFolderBtn");
  const openExistingBtn = document.getElementById("copyProjectLocallyOpenExistingBtn");
  const copyNoticeEl = document.getElementById("copyProjectLocallyCopyNotice");
  const selectDefaultsBtn = document.getElementById(
    "copyProjectLocallySelectDefaultsBtn"
  );
  const selectAllBtn = document.getElementById("copyProjectLocallySelectAllBtn");
  const clearAllBtn = document.getElementById("copyProjectLocallyClearAllBtn");

  if (copyTabBtn) {
    const isActive = activeTab === "copy";
    copyTabBtn.classList.toggle("is-active", isActive);
    copyTabBtn.setAttribute("aria-selected", isActive ? "true" : "false");
    copyTabBtn.tabIndex = isActive ? 0 : -1;
  }
  if (syncTabBtn) {
    const isActive = activeTab === "sync";
    syncTabBtn.classList.toggle("is-active", isActive);
    syncTabBtn.setAttribute("aria-selected", isActive ? "true" : "false");
    syncTabBtn.tabIndex = isActive ? 0 : -1;
  }
  if (copyPanel) copyPanel.hidden = activeTab !== "copy";
  if (syncPanel) syncPanel.hidden = activeTab !== "sync";

  if (projectNameEl) {
    projectNameEl.textContent =
      copyProjectLocallyDialogState.projectName || "Copy to Local";
  }
  if (serverPathEl) {
    serverPathEl.textContent = copyProjectLocallyDialogState.serverProjectPath || "--";
    serverPathEl.title = copyProjectLocallyDialogState.serverProjectPath || "";
  }
  if (localPathEl) {
    localPathEl.textContent = copyProjectLocallyDialogState.localProjectPath || "--";
    localPathEl.title = copyProjectLocallyDialogState.localProjectPath || "";
  }
  const copyPreviewLoaded = copyProjectLocallyDialogState.copyPreviewLoaded === true;
  const localProjectExists = copyProjectLocallyDialogState.localProjectExists === true;
  const hasOptions = copyProjectLocallyDialogState.folderOptions.length > 0;
  const copyReady = copyPreviewLoaded && !localProjectExists && hasOptions;

  if (chooseFolderBtn) {
    chooseFolderBtn.textContent = copyPreviewLoaded
      ? "Change Server Project Folder"
      : "Choose Server Project Folder";
  }
  if (openExistingBtn) {
    const canOpenExisting =
      localProjectExists && !!copyProjectLocallyDialogState.localProjectPath;
    openExistingBtn.hidden = !canOpenExisting;
    openExistingBtn.disabled = !canOpenExisting;
  }
  if (copyNoticeEl) {
    if (copyProjectLocallyDialogState.copyLoadErrorMessage) {
      copyNoticeEl.textContent = copyProjectLocallyDialogState.copyLoadErrorMessage;
    } else if (!copyPreviewLoaded) {
      copyNoticeEl.textContent =
        "Choose a server project folder to review subfolders before copying.";
    } else if (localProjectExists) {
      copyNoticeEl.textContent =
        "A local project copy already exists. Open it here or switch to Sync to Server to review local updates.";
    } else if (!hasOptions) {
      copyNoticeEl.textContent =
        "No subfolders were found in the selected server project folder.";
    } else {
      copyNoticeEl.textContent = "";
    }
  }

  renderCopyProjectLocallyFolderList(listEl, copyReady ? copyProjectLocallyDialogState.folderOptions : []);

  if (selectDefaultsBtn) selectDefaultsBtn.disabled = !copyReady;
  if (selectAllBtn) selectAllBtn.disabled = !copyReady;
  if (clearAllBtn) clearAllBtn.disabled = !copyReady;

  updateCopyProjectLocallyDialogSummary();
  renderLocalProjectManagerSyncSection();
}

async function openCopyProjectLocallyDialog(preview, launchContext = null) {
  const dialog = document.getElementById("copyProjectLocallyDlg");
  if (!dialog) {
    throw new Error("Local Project Manager dialog is unavailable.");
  }

  const previewStatus = String(preview?.status || "").trim().toLowerCase();
  const launchProjectPath = getLaunchContextProjectRoot(launchContext);
  copyProjectLocallyDialogState = {
    ...DEFAULT_COPY_PROJECT_LOCALLY_DIALOG_STATE,
    activeTab: "copy",
    serverProjectPath: String(
      preview?.serverProjectPath || preview?.resolvedServerProjectPath || launchProjectPath || ""
    ).trim(),
    localProjectPath: String(preview?.localProjectPath || "").trim(),
    projectName: String(preview?.projectName || launchContext?.projectName || "").trim(),
    resolvedFromWorkroom: preview?.resolvedFromWorkroom === true,
    resolutionMode: String(preview?.resolutionMode || "").trim(),
    workroomProjectPath: String(preview?.workroomProjectPath || "").trim(),
    copyPreviewLoaded: previewStatus === "success",
    copyLoadErrorMessage:
      preview && previewStatus !== "success"
        ? String(preview?.message || "").trim()
        : "",
    localProjectExists: preview?.localProjectExists === true,
    launchContext: deepCloneJson(launchContext, null),
    folderOptions: Array.isArray(preview?.folderOptions)
      ? preview.folderOptions.map((option, index) =>
          normalizeCopyProjectLocallyFolderOption(option, index)
        )
      : [],
    sync: createDefaultLocalProjectManagerSyncState(),
  };

  renderCopyProjectLocallyDialog();
  queueMicrotask(() => {
    void loadLocalProjectManagerProjects();
  });

  return new Promise((resolve) => {
    const handleClose = () => {
      const action = String(dialog.dataset.localProjectManagerAction || "").trim();
      let result = null;
      if (dialog.returnValue === "confirm") {
        if (action === "sync") {
          result = buildLocalProjectManagerSyncConfirmPayload();
        } else if (action === "copy") {
          result = {
            action: "copy",
            selectionPayload: buildCopyProjectLocallySelectionPayload(),
            serverProjectPath: copyProjectLocallyDialogState.serverProjectPath || "",
            launchContext: deepCloneJson(copyProjectLocallyDialogState.launchContext, null),
          };
        }
      }
      dialog.returnValue = "";
      dialog.dataset.localProjectManagerAction = "";
      resetCopyProjectLocallyDialogState();
      dialog.removeEventListener("close", handleClose);
      resolve(result);
    };

    dialog.dataset.localProjectManagerAction = "";
    dialog.addEventListener("close", handleClose);
    if (!showDialog(dialog)) {
      dialog.removeEventListener("close", handleClose);
      dialog.dataset.localProjectManagerAction = "";
      resetCopyProjectLocallyDialogState();
      resolve(null);
    }
  });
}

function switchLocalProjectManagerTab(nextTab) {
  copyProjectLocallyDialogState.activeTab = nextTab === "sync" ? "sync" : "copy";
  renderCopyProjectLocallyDialog();
}

function normalizeLocalProjectManagerText(value = "") {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function formatLocalProjectManagerTimestampLabel(value = "") {
  const raw = String(value || "").trim();
  if (!raw) return "";
  const parsed = new Date(raw);
  if (!(parsed instanceof Date) || Number.isNaN(parsed.getTime())) return raw;
  return parsed.toLocaleString([], {
    year: "numeric",
    month: "short",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit",
  });
}

function extractLocalProjectManagerProjectId(rawPath = "") {
  const normalizedPath = normalizeWindowsPath(rawPath);
  const parsed = parseProjectFromPath(normalizedPath || rawPath);
  if (parsed?.id) return String(parsed.id || "").trim();
  const leaf = getWindowsPathLeaf(normalizedPath || String(rawPath || "").trim());
  const match = String(leaf || "").match(/(?:^|[^0-9])(\d{6})(?!\d)/);
  return match ? String(match[1] || "").trim() : "";
}

function compareLocalProjectManagerProjectIdentity(projectA, projectB) {
  const idDiff = String(projectA?.id || "").localeCompare(
    String(projectB?.id || ""),
    undefined,
    { numeric: true, sensitivity: "base" }
  );
  if (idDiff) return idDiff;
  const nameDiff = String(projectA?.name || "").localeCompare(
    String(projectB?.name || ""),
    undefined,
    { numeric: true, sensitivity: "base" }
  );
  if (nameDiff) return nameDiff;
  return String(projectA?.nick || "").localeCompare(
    String(projectB?.nick || ""),
    undefined,
    { numeric: true, sensitivity: "base" }
  );
}

function getLocalProjectManagerProjectSortMeta(project) {
  const priority = getProjectListPriorityMeta(project);
  return {
    bucket: Number.isFinite(priority?.sortBucket) ? priority.sortBucket : 2,
    activeDue: priority?.sortDueDate || null,
    fallbackDue: priority?.fallbackDueDate || null,
  };
}

function compareLocalProjectManagerProjects(projectA, projectB) {
  const metaA = getLocalProjectManagerProjectSortMeta(projectA);
  const metaB = getLocalProjectManagerProjectSortMeta(projectB);
  if (metaA.bucket !== metaB.bucket) return metaA.bucket - metaB.bucket;

  if (metaA.bucket === 0) {
    const dueDiff = compareDueDateValues(metaA.activeDue, metaB.activeDue, "asc");
    if (dueDiff) return dueDiff;
  }

  if (metaA.bucket === 2) {
    const completedDiff = compareDueDateValues(
      metaA.fallbackDue,
      metaB.fallbackDue,
      "asc"
    );
    if (completedDiff) return completedDiff;
  }

  return compareLocalProjectManagerProjectIdentity(projectA, projectB);
}

function getLocalProjectManagerSortedProjects() {
  return (Array.isArray(db) ? db.slice() : []).filter(Boolean).sort(compareLocalProjectManagerProjects);
}

function findLocalProjectManagerMatchedProject(localProjectPath, folderName, sortedProjects) {
  const normalizedLocalProjectPath = normalizeWindowsPath(localProjectPath).toLowerCase();
  if (normalizedLocalProjectPath) {
    const exactMatch = sortedProjects.find(
      (project) =>
        normalizeWindowsPath(project?.localProjectPath || "").toLowerCase() ===
        normalizedLocalProjectPath
    );
    if (exactMatch) return exactMatch;
  }

  const extractedProjectId = extractLocalProjectManagerProjectId(
    localProjectPath || folderName
  );
  if (extractedProjectId) {
    const idMatch = sortedProjects.find(
      (project) => String(project?.id || "").trim() === extractedProjectId
    );
    if (idMatch) return idMatch;
  }

  const folderKey = normalizeLocalProjectManagerText(
    folderName || getWindowsPathLeaf(localProjectPath)
  );
  if (!folderKey) return null;

  return (
    sortedProjects.find((project) => {
      const projectName = normalizeLocalProjectManagerText(project?.name || "");
      const projectNick = normalizeLocalProjectManagerText(project?.nick || "");
      const projectPathLeaf = normalizeLocalProjectManagerText(
        getWindowsPathLeaf(normalizeProjectPath(project?.path || ""))
      );
      const localPathLeaf = normalizeLocalProjectManagerText(
        getWindowsPathLeaf(normalizeWindowsPath(project?.localProjectPath || ""))
      );
      return [projectName, projectNick, projectPathLeaf, localPathLeaf].includes(
        folderKey
      );
    }) || null
  );
}

function buildLocalProjectManagerProjectEntry(rawEntry, sortedProjects, index = 0) {
  const localProjectPath = normalizeWindowsPath(rawEntry?.localProjectPath || "");
  const name =
    String(rawEntry?.name || "").trim() ||
    getWindowsPathLeaf(localProjectPath) ||
    `Local Project ${index + 1}`;
  const matchedProject = findLocalProjectManagerMatchedProject(
    localProjectPath,
    name,
    sortedProjects
  );
  const matchedProjectIndex = matchedProject ? db.indexOf(matchedProject) : -1;
  const matchedProjectId = String(
    matchedProject?.id || extractLocalProjectManagerProjectId(localProjectPath || name) || ""
  ).trim();
  const matchedProjectName = String(matchedProject?.name || "").trim();
  const matchedProjectNick = String(matchedProject?.nick || "").trim();
  return {
    id: `local-project-manager::${localProjectPath.toLowerCase() || index}`,
    name,
    displayName: matchedProjectName || name,
    localProjectPath,
    lastModifiedAt: String(rawEntry?.lastModifiedAt || "").trim(),
    lastModifiedLabel:
      String(rawEntry?.lastModifiedLabel || "").trim() ||
      formatLocalProjectManagerTimestampLabel(rawEntry?.lastModifiedAt || ""),
    matchedProject,
    matchedProjectIndex,
    matchedProjectId,
    matchedProjectName,
    matchedProjectNick,
    matchedServerProjectPath: matchedProject
      ? getWorkroomServerProjectPath(matchedProject) ||
        normalizeProjectPath(matchedProject?.path || "")
      : "",
    searchText: normalizeLocalProjectManagerText(
      [
        name,
        localProjectPath,
        matchedProjectId,
        matchedProjectName,
        matchedProjectNick,
      ].join(" ")
    ),
  };
}

function compareLocalProjectManagerProjectEntries(a, b) {
  const aMatched = Number.isInteger(a?.matchedProjectIndex) && a.matchedProjectIndex >= 0;
  const bMatched = Number.isInteger(b?.matchedProjectIndex) && b.matchedProjectIndex >= 0;
  if (aMatched !== bMatched) return aMatched ? -1 : 1;
  if (aMatched && bMatched) {
    const matchedDiff = compareLocalProjectManagerProjects(
      a.matchedProject,
      b.matchedProject
    );
    if (matchedDiff) return matchedDiff;
  }
  return String(a?.name || a?.localProjectPath || "").localeCompare(
    String(b?.name || b?.localProjectPath || ""),
    undefined,
    { numeric: true, sensitivity: "base" }
  );
}

function normalizeLocalProjectManagerSyncCandidateFile(candidate, index = 0) {
  const sizeBytesValue = Number(candidate?.sizeBytes);
  return {
    id: `local-project-manager-sync::${candidate?.relativePath || index}`,
    relativePath: String(candidate?.relativePath || "").trim(),
    localPath: normalizeWindowsPath(candidate?.localPath || ""),
    serverPath: normalizeWindowsPath(candidate?.serverPath || ""),
    reason:
      String(candidate?.reason || "").trim().toLowerCase() === "server_missing"
        ? "server_missing"
        : "local_newer",
    localModifiedAt: String(candidate?.localModifiedAt || "").trim(),
    serverModifiedAt: String(candidate?.serverModifiedAt || "").trim(),
    sizeBytes: Number.isFinite(sizeBytesValue) ? sizeBytesValue : null,
    sizeLabel:
      String(candidate?.sizeLabel || "").trim() ||
      (Number.isFinite(sizeBytesValue)
        ? formatCopyProjectLocallySizeLabel(sizeBytesValue)
        : "Size unavailable"),
    selected: candidate?.selectedByDefault === true,
  };
}

function normalizeLocalProjectManagerBlockedEntry(entry, index = 0) {
  return {
    id: `local-project-manager-blocked::${entry?.relativePath || index}`,
    relativePath: String(entry?.relativePath || "").trim(),
    localPath: normalizeWindowsPath(entry?.localPath || ""),
    serverPath: normalizeWindowsPath(entry?.serverPath || ""),
    reason: String(entry?.reason || "").trim() || "blocked",
    error: String(entry?.error || "").trim(),
  };
}

function resetLocalProjectManagerSyncPreview({ keepManualServerPath = true } = {}) {
  const syncState = copyProjectLocallyDialogState.sync;
  syncState.previewLoading = false;
  syncState.previewLoaded = false;
  syncState.previewError = "";
  syncState.candidateFiles = [];
  syncState.blockedEntries = [];
  syncState.resolvedServerProjectPath = "";
  syncState.serverPathSource = "";
  if (!keepManualServerPath) {
    syncState.manualServerProjectPath = "";
  }
}

function getFilteredLocalProjectManagerProjectEntries() {
  const syncState = copyProjectLocallyDialogState.sync;
  const query = normalizeLocalProjectManagerText(syncState.searchQuery || "");
  const items = Array.isArray(syncState.projectEntries) ? syncState.projectEntries : [];
  if (!query) return items;
  const queryTokens = query.split(/\s+/).filter(Boolean);
  return items.filter((entry) =>
    queryTokens.every((token) => String(entry?.searchText || "").includes(token))
  );
}

function selectLocalProjectManagerProjectEntry(entry) {
  if (!entry) return;
  const syncState = copyProjectLocallyDialogState.sync;
  syncState.selectedLocalProjectPath = entry.localProjectPath || "";
  syncState.localProjectPath = entry.localProjectPath || "";
  syncState.localProjectName = entry.displayName || entry.name || "Local Project";
  syncState.manualLocalProjectPath = entry.localProjectPath || "";
  syncState.matchedProjectIndex = Number.isInteger(entry.matchedProjectIndex)
    ? entry.matchedProjectIndex
    : -1;
  syncState.matchedProjectId = entry.matchedProjectId || "";
  syncState.matchedProjectName = entry.matchedProjectName || "";
  syncState.matchedProjectNick = entry.matchedProjectNick || "";
  resetLocalProjectManagerSyncPreview({ keepManualServerPath: false });
  renderCopyProjectLocallyDialog();
}

function createManualLocalProjectManagerProjectEntry(pathValue = "") {
  const normalizedPath = normalizeWindowsPath(pathValue);
  const sortedProjects = getLocalProjectManagerSortedProjects();
  return buildLocalProjectManagerProjectEntry(
    {
      name: getWindowsPathLeaf(normalizedPath) || normalizedPath,
      localProjectPath: normalizedPath,
      lastModifiedAt: "",
      lastModifiedLabel: "",
    },
    sortedProjects,
    Date.now()
  );
}

function getLocalProjectManagerSyncServerPathInfo() {
  const syncState = copyProjectLocallyDialogState.sync;
  const manualServerPath = normalizeWindowsPath(syncState.manualServerProjectPath || "");
  if (manualServerPath) {
    return { path: manualServerPath, source: "manual" };
  }

  const launchSource = String(copyProjectLocallyDialogState.launchContext?.source || "")
    .trim()
    .toLowerCase();
  const launchProjectPath = getLaunchContextProjectRoot(
    copyProjectLocallyDialogState.launchContext
  );
  if (launchSource === "workroom") {
    const activeProject = db[activeChecklistProject] || null;
    const workroomPath =
      getWorkroomServerProjectPath(activeProject) ||
      normalizeProjectPath(
        copyProjectLocallyDialogState.launchContext?.rootProjectPath ||
          copyProjectLocallyDialogState.launchContext?.projectPath ||
          ""
      );
    if (workroomPath) return { path: workroomPath, source: "workroom" };
  }
  if (launchProjectPath) {
    return { path: launchProjectPath, source: launchSource || "launch_context" };
  }

  const matchedProject =
    Number.isInteger(syncState.matchedProjectIndex) &&
    syncState.matchedProjectIndex >= 0
      ? db[syncState.matchedProjectIndex] || null
      : null;
  if (matchedProject) {
    const matchedPath =
      getWorkroomServerProjectPath(matchedProject) ||
      normalizeProjectPath(matchedProject?.path || "");
    if (matchedPath) return { path: matchedPath, source: "matched_project" };
  }

  return { path: "", source: "" };
}

function createLocalProjectManagerProjectRow(entry) {
  const syncState = copyProjectLocallyDialogState.sync;
  const isSelected =
    normalizeWindowsPath(syncState.localProjectPath || "").toLowerCase() ===
    normalizeWindowsPath(entry?.localProjectPath || "").toLowerCase();
  const matchedMeta = entry?.matchedProjectId
    ? `${entry.matchedProjectId}${entry.matchedProjectName ? ` • ${entry.matchedProjectName}` : ""}${
        entry.matchedProjectNick ? ` (${entry.matchedProjectNick})` : ""
      }`
    : "Unmatched local project";
  const row = el(
    "button",
    {
      type: "button",
      className: "local-project-manager-project-row",
      "data-local-project-path": entry?.localProjectPath || "",
    },
    [
      el("div", { className: "local-project-manager-project-title", textContent: entry?.displayName || entry?.name || "Local Project" }),
      el("div", { className: "local-project-manager-project-meta", textContent: matchedMeta }),
      el("div", { className: "local-project-manager-project-path", textContent: entry?.localProjectPath || "--" }),
      el("div", { className: "local-project-manager-project-modified", textContent: entry?.lastModifiedLabel ? `Updated ${entry.lastModifiedLabel}` : "" }),
    ]
  );
  row.classList.toggle("is-selected", isSelected);
  row.setAttribute("aria-pressed", isSelected ? "true" : "false");
  return row;
}

function renderLocalProjectManagerSyncProjectList(container, items) {
  if (!container) return;
  container.replaceChildren();

  const syncState = copyProjectLocallyDialogState.sync;
  if (syncState.loadingProjects) {
    container.appendChild(
      el("div", {
        className: "deliverable-notepad-empty",
        textContent: "Loading local projects...",
      })
    );
    return;
  }

  if (syncState.projectsError) {
    container.appendChild(
      el("div", {
        className: "deliverable-notepad-empty",
        textContent: syncState.projectsError,
      })
    );
    return;
  }

  if (!items.length) {
    container.appendChild(
      el("div", {
        className: "deliverable-notepad-empty",
        textContent: "No local projects matched the current search.",
      })
    );
    return;
  }

  items.forEach((entry) => container.appendChild(createLocalProjectManagerProjectRow(entry)));
}

function createLocalProjectManagerSyncCandidateRow(candidate) {
  const reasonLabel =
    candidate.reason === "server_missing" ? "Missing on server" : "Local file is newer";
  const localModifiedLabel = formatLocalProjectManagerTimestampLabel(
    candidate.localModifiedAt || ""
  );
  const serverModifiedLabel = formatLocalProjectManagerTimestampLabel(
    candidate.serverModifiedAt || ""
  );
  const metaParts = [reasonLabel];
  if (localModifiedLabel) metaParts.push(`Local ${localModifiedLabel}`);
  if (serverModifiedLabel) metaParts.push(`Server ${serverModifiedLabel}`);

  const row = el(
    "label",
    {
      className: "copy-project-locally-row local-project-manager-sync-row",
      role: "listitem",
    },
    [
      el("span", { className: "custom-check copy-project-locally-checkbox" }, [
        el("input", {
          className: "copy-project-locally-checkbox-input",
          type: "checkbox",
          checked: candidate.selected === true,
          "data-local-project-manager-sync-checkbox": "true",
          "data-relative-path": candidate.relativePath || "",
          "aria-label": `Sync ${candidate.relativePath || "file"} to server`,
        }),
        el("span", {
          className: "checkmark",
          "aria-hidden": "true",
        }),
      ]),
      el("div", { className: "local-project-manager-sync-row-content" }, [
        el("div", {
          className: "copy-project-locally-row-title",
          textContent: candidate.relativePath || "Untitled file",
        }),
        el("div", {
          className: "local-project-manager-sync-row-meta",
          textContent: metaParts.join(" • "),
        }),
      ]),
      el("div", {
        className: "copy-project-locally-row-size",
        textContent: candidate.sizeLabel || "Size unavailable",
      }),
    ]
  );

  syncCopyProjectLocallyRowSelectionState(row, candidate.selected === true, "sync");
  return row;
}

function renderLocalProjectManagerSyncRecommendations(container, items) {
  if (!container) return;
  container.replaceChildren();

  const syncState = copyProjectLocallyDialogState.sync;
  if (syncState.previewLoading) {
    container.appendChild(
      el("div", {
        className: "deliverable-notepad-empty",
        textContent: "Loading sync recommendations...",
      })
    );
    return;
  }

  if (syncState.previewError) {
    container.appendChild(
      el("div", {
        className: "deliverable-notepad-empty",
        textContent: syncState.previewError,
      })
    );
    return;
  }

  if (!syncState.previewLoaded) {
    container.appendChild(
      el("div", {
        className: "deliverable-notepad-empty",
        textContent: "Load recommendations to review newer local files.",
      })
    );
    return;
  }

  if (!items.length) {
    container.appendChild(
      el("div", {
        className: "deliverable-notepad-empty",
        textContent: "No newer local files were found.",
      })
    );
    return;
  }

  items.forEach((item) => container.appendChild(createLocalProjectManagerSyncCandidateRow(item)));
}

function renderLocalProjectManagerSyncBlockedEntries(container, items) {
  const blockedSection = document.getElementById("localProjectManagerSyncBlockedSection");
  if (blockedSection) blockedSection.hidden = !items.length;
  if (!container) return;
  container.replaceChildren();
  items.forEach((entry) => {
    const labelParts = [];
    if (entry.relativePath) labelParts.push(entry.relativePath);
    if (entry.reason) labelParts.push(entry.reason.replace(/_/g, " "));
    if (entry.error) labelParts.push(entry.error);
    container.appendChild(
      el("div", {
        className: "local-project-manager-blocked-entry",
        textContent: labelParts.join(" • "),
      })
    );
  });
}

function buildLocalProjectManagerSyncSelectionPayload() {
  const syncState = copyProjectLocallyDialogState.sync;
  const selectedFiles = syncState.candidateFiles.filter((entry) => entry.selected === true);
  return {
    hasSelection: selectedFiles.length > 0,
    selectedRelativePaths: selectedFiles
      .map((entry) => entry.relativePath)
      .filter(Boolean),
  };
}

function buildLocalProjectManagerSyncConfirmPayload() {
  const syncState = copyProjectLocallyDialogState.sync;
  const selectionPayload = buildLocalProjectManagerSyncSelectionPayload();
  const serverPathInfo = getLocalProjectManagerSyncServerPathInfo();
  if (
    syncState.previewLoaded !== true ||
    !selectionPayload.hasSelection ||
    !syncState.localProjectPath
  ) {
    return null;
  }
  return {
    action: "sync",
    localProjectPath: syncState.localProjectPath || "",
    serverProjectPath:
      syncState.resolvedServerProjectPath || serverPathInfo.path || "",
    selectedRelativePaths: selectionPayload.selectedRelativePaths,
    launchContext: deepCloneJson(copyProjectLocallyDialogState.launchContext, null),
  };
}

function updateLocalProjectManagerSyncSummary() {
  const summaryEl = document.getElementById("localProjectManagerSyncSummary");
  const confirmBtn = document.getElementById("localProjectManagerSyncConfirmBtn");
  const selectAllBtn = document.getElementById("localProjectManagerSyncSelectAllBtn");
  const clearAllBtn = document.getElementById("localProjectManagerSyncClearAllBtn");
  const previewBtn = document.getElementById("localProjectManagerPreviewSyncBtn");
  const syncState = copyProjectLocallyDialogState.sync;
  const selectionPayload = buildLocalProjectManagerSyncSelectionPayload();
  const hasCandidates = syncState.candidateFiles.length > 0;

  if (confirmBtn) {
    confirmBtn.disabled =
      syncState.previewLoaded !== true || selectionPayload.hasSelection !== true;
  }
  if (selectAllBtn) selectAllBtn.disabled = syncState.previewLoaded !== true || !hasCandidates;
  if (clearAllBtn) clearAllBtn.disabled = syncState.previewLoaded !== true || !hasCandidates;
  if (previewBtn) previewBtn.disabled = syncState.previewLoading === true || !syncState.localProjectPath;

  if (!summaryEl) return;
  if (!syncState.localProjectPath) {
    summaryEl.textContent = "Select a local project to review newer files.";
    return;
  }
  if (syncState.previewLoading) {
    summaryEl.textContent = "Loading recommendations...";
    return;
  }
  if (syncState.previewError) {
    summaryEl.textContent = syncState.previewError;
    return;
  }
  if (syncState.previewLoaded !== true) {
    summaryEl.textContent = "Load recommendations to review local updates.";
    return;
  }
  if (!hasCandidates) {
    summaryEl.textContent = syncState.blockedEntries.length
      ? "No sync recommendations. Review blocked entries below."
      : "No newer local files were found.";
    return;
  }

  const selectedFiles = syncState.candidateFiles.filter((entry) => entry.selected === true);
  const totalBytes = selectedFiles.reduce((sum, entry) => {
    return sum + (Number.isFinite(entry?.sizeBytes) ? Number(entry.sizeBytes) : 0);
  }, 0);
  summaryEl.textContent = `${selectedFiles.length} file${
    selectedFiles.length === 1 ? "" : "s"
  } selected - ${formatCopyProjectLocallySizeLabel(totalBytes)}`;
}

function renderLocalProjectManagerSyncSection() {
  const syncState = copyProjectLocallyDialogState.sync;
  const searchInput = document.getElementById("localProjectManagerSyncSearch");
  const manualLocalPathInput = document.getElementById("localProjectManagerManualLocalPath");
  const manualServerPathInput = document.getElementById("localProjectManagerManualServerPath");
  const projectListEl = document.getElementById("localProjectManagerSyncProjectList");
  const projectNameEl = document.getElementById("localProjectManagerSyncProjectName");
  const localPathEl = document.getElementById("localProjectManagerSyncLocalPath");
  const serverPathEl = document.getElementById("localProjectManagerSyncServerPath");
  const recommendationListEl = document.getElementById(
    "localProjectManagerSyncRecommendationList"
  );
  const blockedListEl = document.getElementById("localProjectManagerSyncBlockedList");

  if (searchInput) searchInput.value = syncState.searchQuery || "";
  if (manualLocalPathInput) {
    manualLocalPathInput.value = syncState.manualLocalProjectPath || "";
    manualLocalPathInput.title = syncState.manualLocalProjectPath || "";
  }
  if (manualServerPathInput) {
    manualServerPathInput.value = syncState.manualServerProjectPath || "";
    manualServerPathInput.title = syncState.manualServerProjectPath || "";
  }

  const serverPathInfo = getLocalProjectManagerSyncServerPathInfo();
  const displayedServerPath =
    syncState.resolvedServerProjectPath || serverPathInfo.path || "--";
  if (projectNameEl) {
    projectNameEl.textContent = syncState.localProjectName || "Sync to Server";
  }
  if (localPathEl) {
    localPathEl.textContent = syncState.localProjectPath || "--";
    localPathEl.title = syncState.localProjectPath || "";
  }
  if (serverPathEl) {
    serverPathEl.textContent = displayedServerPath;
    serverPathEl.title = displayedServerPath === "--" ? "" : displayedServerPath;
  }

  renderLocalProjectManagerSyncProjectList(
    projectListEl,
    getFilteredLocalProjectManagerProjectEntries()
  );
  renderLocalProjectManagerSyncRecommendations(
    recommendationListEl,
    syncState.candidateFiles
  );
  renderLocalProjectManagerSyncBlockedEntries(blockedListEl, syncState.blockedEntries);
  updateLocalProjectManagerSyncSummary();
}

async function loadLocalProjectManagerProjects() {
  const syncState = copyProjectLocallyDialogState.sync;
  syncState.loadingProjects = true;
  syncState.projectsLoaded = false;
  syncState.projectsError = "";
  renderCopyProjectLocallyDialog();

  try {
    const result = await window.pywebview.api.list_local_project_manager_projects();
    if (result?.status !== "success") {
      throw new Error(result?.message || "Failed to load local projects.");
    }

    const sortedProjects = getLocalProjectManagerSortedProjects();
    const entries = Array.isArray(result?.projects)
      ? result.projects.map((entry, index) =>
          buildLocalProjectManagerProjectEntry(entry, sortedProjects, index)
        )
      : [];
    entries.sort(compareLocalProjectManagerProjectEntries);

    syncState.localRootPath = String(result?.localRootPath || "").trim();
    syncState.projectEntries = entries;
    syncState.projectsLoaded = true;
  } catch (error) {
    syncState.projectsError = error?.message || "Failed to load local projects.";
  } finally {
    syncState.loadingProjects = false;
    renderCopyProjectLocallyDialog();
  }
}

function applyLocalProjectManagerManualLocalPath() {
  const manualLocalPathInput = document.getElementById("localProjectManagerManualLocalPath");
  const normalizedPath = normalizeWindowsPath(manualLocalPathInput?.value || "");
  if (!normalizedPath) {
    toast("Choose a local project folder first.");
    return;
  }

  const existingEntry = copyProjectLocallyDialogState.sync.projectEntries.find(
    (entry) =>
      normalizeWindowsPath(entry?.localProjectPath || "").toLowerCase() ===
      normalizedPath.toLowerCase()
  );
  selectLocalProjectManagerProjectEntry(
    existingEntry || createManualLocalProjectManagerProjectEntry(normalizedPath)
  );
}

async function browseLocalProjectManagerLocalPath() {
  const selection = await window.pywebview.api.select_folder();
  if (selection?.status === "error") {
    throw new Error(selection.message || "Failed to choose a local project folder.");
  }
  if (!selection || selection.status === "cancelled" || !selection.path) {
    return;
  }
  const syncState = copyProjectLocallyDialogState.sync;
  syncState.manualLocalProjectPath = normalizeWindowsPath(selection.path || "");
  renderCopyProjectLocallyDialog();
}

async function browseLocalProjectManagerServerPath() {
  const selection = await window.pywebview.api.select_folder(
    getLaunchContextProjectRoot(copyProjectLocallyDialogState.launchContext) || null
  );
  if (selection?.status === "error") {
    throw new Error(selection.message || "Failed to choose a server project folder.");
  }
  if (!selection || selection.status === "cancelled" || !selection.path) {
    return;
  }
  const syncState = copyProjectLocallyDialogState.sync;
  syncState.manualServerProjectPath = normalizeWindowsPath(selection.path || "");
  resetLocalProjectManagerSyncPreview();
  renderCopyProjectLocallyDialog();
}

function applyCopyProjectLocallyPreviewResult(preview, launchContext = null) {
  const previewStatus = String(preview?.status || "").trim().toLowerCase();
  copyProjectLocallyDialogState.serverProjectPath = String(
    preview?.serverProjectPath || preview?.resolvedServerProjectPath || ""
  ).trim();
  copyProjectLocallyDialogState.localProjectPath = String(preview?.localProjectPath || "").trim();
  copyProjectLocallyDialogState.projectName = String(preview?.projectName || "").trim();
  copyProjectLocallyDialogState.resolvedFromWorkroom = preview?.resolvedFromWorkroom === true;
  copyProjectLocallyDialogState.resolutionMode = String(preview?.resolutionMode || "").trim();
  copyProjectLocallyDialogState.workroomProjectPath = String(preview?.workroomProjectPath || "").trim();
  copyProjectLocallyDialogState.launchContext = deepCloneJson(
    launchContext ?? copyProjectLocallyDialogState.launchContext,
    null
  );
  copyProjectLocallyDialogState.copyPreviewLoaded = previewStatus === "success";
  copyProjectLocallyDialogState.copyLoadErrorMessage =
    previewStatus === "success" ? "" : String(preview?.message || "").trim();
  copyProjectLocallyDialogState.localProjectExists = preview?.localProjectExists === true;
  copyProjectLocallyDialogState.folderOptions =
    previewStatus === "success" && Array.isArray(preview?.folderOptions)
      ? preview.folderOptions.map((option, index) =>
          normalizeCopyProjectLocallyFolderOption(option, index)
        )
      : [];
}

async function chooseAndPreviewCopyProjectLocallyServerFolder() {
  const selection = await window.pywebview.api.select_folder(
    getLaunchContextProjectRoot(copyProjectLocallyDialogState.launchContext) || null
  );
  if (selection?.status === "error") {
    throw new Error(selection.message || "Failed to choose a folder.");
  }
  if (!selection || selection.status === "cancelled" || !selection.path) {
    return;
  }

  const launchContext = copyProjectLocallyDialogState.launchContext || null;
  const preview = await window.pywebview.api.preview_copy_project_locally(
    String(selection.path || "").trim(),
    launchContext
  );
  if (preview?.status !== "success") {
    throw new Error(preview?.message || "Failed to load project folders.");
  }
  applyCopyProjectLocallyPreviewResult(preview, launchContext);
  renderCopyProjectLocallyDialog();
}

async function previewLocalProjectManagerSync() {
  const syncState = copyProjectLocallyDialogState.sync;
  if (!syncState.localProjectPath) {
    toast("Select a local project first.");
    return;
  }

  const serverPathInfo = getLocalProjectManagerSyncServerPathInfo();
  if (!serverPathInfo.path) {
    toast("Choose a server project folder before loading recommendations.", 7000);
    return;
  }

  syncState.previewLoading = true;
  syncState.previewLoaded = false;
  syncState.previewError = "";
  syncState.candidateFiles = [];
  syncState.blockedEntries = [];
  renderCopyProjectLocallyDialog();

  try {
    const result = await window.pywebview.api.preview_local_project_manager_sync(
      syncState.localProjectPath,
      serverPathInfo.path,
      copyProjectLocallyDialogState.launchContext || null
    );
    if (result?.status !== "success") {
      throw new Error(result?.message || "Failed to load sync recommendations.");
    }

    syncState.resolvedServerProjectPath = String(
      result?.serverProjectPath || result?.resolvedServerProjectPath || serverPathInfo.path || ""
    ).trim();
    syncState.serverPathSource = serverPathInfo.source || "";
    syncState.candidateFiles = Array.isArray(result?.candidateFiles)
      ? result.candidateFiles.map((candidate, index) =>
          normalizeLocalProjectManagerSyncCandidateFile(candidate, index)
        )
      : [];
    syncState.blockedEntries = Array.isArray(result?.blockedEntries)
      ? result.blockedEntries.map((entry, index) =>
          normalizeLocalProjectManagerBlockedEntry(entry, index)
        )
      : [];
    syncState.previewLoaded = true;
  } catch (error) {
    syncState.previewError =
      error?.message || "Failed to load sync recommendations.";
  } finally {
    syncState.previewLoading = false;
    renderCopyProjectLocallyDialog();
  }
}

let deliverableNotepadEntries = [];
let deliverableNotepadSelectedEntryIds = [];

function formatDeliverableExportField(value, fallback = "--") {
  const normalized = String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();
  return normalized || fallback;
}

function getDeliverableStatusText(deliverable) {
  const statuses = Array.isArray(deliverable?.statuses)
    ? deliverable.statuses
      .map((status) => String(status || "").trim())
      .filter(Boolean)
    : [];
  if (statuses.length) return [...new Set(statuses)].join(", ");

  const singleStatus = String(deliverable?.status || "").trim();
  return singleStatus || "None";
}

function hasDeliverableExportContent(deliverable) {
  if (!deliverable || typeof deliverable !== "object") return false;
  if (String(deliverable.name || "").trim()) return true;
  if (String(deliverable.due || "").trim()) return true;
  if (String(deliverable.notes || "").trim()) return true;
  if (String(deliverable.status || "").trim()) return true;
  if (
    Array.isArray(deliverable.statuses) &&
    deliverable.statuses.some((status) => String(status || "").trim())
  ) {
    return true;
  }
  if (
    Array.isArray(deliverable.tasks) &&
    deliverable.tasks.some((task) => String(task?.text || "").trim())
  ) {
    return true;
  }
  return false;
}

function buildDeliverableNotepadEntries(projects = db) {
  const candidates = Array.isArray(projects) ? projects : [];
  return candidates
    .flatMap((project, projectIndex) =>
      getProjectDeliverables(project)
        .filter((deliverable) => hasDeliverableExportContent(deliverable))
        .map((deliverable, deliverableIndex) => {
          const projectId = String(project?.id || "").trim();
          const projectName = String(project?.name || "").trim();
          const deliverableId = String(deliverable?.id || "").trim();
          const deliverableName = String(deliverable?.name || "").trim();
          return {
            id: `${projectId || projectName || `project-${projectIndex}`}::${
              deliverableId ||
              deliverableName ||
              `deliverable-${deliverableIndex}`
            }`,
            projectId,
            projectName,
            deliverableName,
            due: String(deliverable?.due || "").trim(),
            dueLabel: humanDate(deliverable?.due) || "No date",
            statusText: getDeliverableStatusText(deliverable),
            project,
            deliverable,
          };
        })
    )
    .sort((a, b) => {
      const dueCompare = compareDeliverablesByDueDesc(a.deliverable, b.deliverable);
      if (dueCompare) return dueCompare;
      const projectCompare = String(a.projectId || a.projectName || "").localeCompare(
        String(b.projectId || b.projectName || ""),
        undefined,
        { numeric: true, sensitivity: "base" }
      );
      if (projectCompare) return projectCompare;
      return String(a.deliverableName || "").localeCompare(
        String(b.deliverableName || ""),
        undefined,
        { numeric: true, sensitivity: "base" }
      );
    });
}

function buildSelectedDeliverablesExcelRows(entries = []) {
  const selectedEntries = Array.isArray(entries) ? entries.filter(Boolean) : [];
  const sortedEntries = selectedEntries.slice().sort((a, b) => {
    const dueCompare = compareDeliverablesByDueDesc(a?.deliverable, b?.deliverable);
    if (dueCompare) return dueCompare;
    const projectCompare = String(a?.projectId || a?.projectName || "").localeCompare(
      String(b?.projectId || b?.projectName || ""),
      undefined,
      { numeric: true, sensitivity: "base" }
    );
    if (projectCompare) return projectCompare;
    return String(a?.deliverableName || "").localeCompare(
      String(b?.deliverableName || ""),
      undefined,
      { numeric: true, sensitivity: "base" }
    );
  });

  return {
    entries: sortedEntries.map((item) => ({
      projectId: String(item?.projectId || "").trim(),
      projectName: String(item?.projectName || "").trim(),
      deliverableName:
        String(item?.deliverableName || "").trim() || "Untitled Deliverable",
      due: String(item?.due || "").trim(),
      statusText: String(item?.statusText || "").trim() || "None",
    })),
    deliverableCount: sortedEntries.length,
  };
}

function createDeliverableNotepadListItem(item) {
  return el("label", { className: "deliverable-notepad-item", role: "listitem" }, [
    el("input", {
      type: "checkbox",
      value: item.id,
      "aria-label": `Select ${item.deliverableName || "untitled deliverable"}`,
    }),
    el("div", { className: "deliverable-notepad-item-content" }, [
      el("div", {
        className: "deliverable-notepad-item-title",
        textContent: item.deliverableName || "Untitled Deliverable",
      }),
      el("div", {
        className: "deliverable-notepad-item-subtitle",
        textContent: `Project ID: ${formatDeliverableExportField(item.projectId)}`,
      }),
      el("div", {
        className: "deliverable-notepad-item-subtitle",
        textContent: `Project Name: ${formatDeliverableExportField(
          item.projectName
        )}`,
      }),
      el("div", {
        className: "deliverable-notepad-item-meta",
        textContent: `Due: ${item.dueLabel} | Status: ${item.statusText}`,
      }),
    ]),
  ]);
}

function renderDeliverableNotepadList(container, items, emptyMessage) {
  if (!container) return;
  container.replaceChildren();
  if (!items.length) {
    container.appendChild(
      el("div", {
        className: "deliverable-notepad-empty",
        textContent: emptyMessage,
      })
    );
    return;
  }
  items.forEach((item) => container.appendChild(createDeliverableNotepadListItem(item)));
}

function renderDeliverableNotepadDialog() {
  const availableList = document.getElementById("deliverableNotepadAvailableList");
  const selectedList = document.getElementById("deliverableNotepadSelectedList");
  if (!availableList || !selectedList) return;

  const entryMap = new Map(deliverableNotepadEntries.map((item) => [item.id, item]));
  const selectedIdSet = new Set(deliverableNotepadSelectedEntryIds);
  const availableItems = deliverableNotepadEntries.filter(
    (item) => !selectedIdSet.has(item.id)
  );
  const selectedItems = deliverableNotepadSelectedEntryIds
    .map((id) => entryMap.get(id))
    .filter(Boolean);

  renderDeliverableNotepadList(
    availableList,
    availableItems,
    "No deliverables are available to add."
  );
  renderDeliverableNotepadList(
    selectedList,
    selectedItems,
    "No deliverables have been selected for export."
  );

  const availableCount = document.getElementById("deliverableNotepadAvailableCount");
  if (availableCount) {
    availableCount.textContent = `${availableItems.length} deliverable${
      availableItems.length === 1 ? "" : "s"
    }`;
  }

  const selectedCount = document.getElementById("deliverableNotepadSelectedCount");
  if (selectedCount) {
    selectedCount.textContent = `${selectedItems.length} deliverable${
      selectedItems.length === 1 ? "" : "s"
    }`;
  }

  const addBtn = document.getElementById("deliverableNotepadAddBtn");
  if (addBtn) addBtn.disabled = !availableItems.length;

  const removeBtn = document.getElementById("deliverableNotepadRemoveBtn");
  if (removeBtn) removeBtn.disabled = !selectedItems.length;

  const exportBtn = document.getElementById("deliverableNotepadExportBtn");
  if (exportBtn) exportBtn.disabled = !selectedItems.length;
}

function getCheckedDeliverableNotepadEntryIds(listId) {
  const list = document.getElementById(listId);
  if (!list) return [];
  return [...list.querySelectorAll('input[type="checkbox"]:checked')].map(
    (input) => input.value
  );
}

function addDeliverablesToNotepadSelection() {
  const ids = getCheckedDeliverableNotepadEntryIds("deliverableNotepadAvailableList");
  if (!ids.length) {
    toast("Select at least one deliverable to add.");
    return;
  }

  const selectedIdSet = new Set(deliverableNotepadSelectedEntryIds);
  ids.forEach((id) => {
    if (!selectedIdSet.has(id)) {
      deliverableNotepadSelectedEntryIds.push(id);
      selectedIdSet.add(id);
    }
  });
  renderDeliverableNotepadDialog();
}

function removeDeliverablesFromNotepadSelection() {
  const ids = new Set(
    getCheckedDeliverableNotepadEntryIds("deliverableNotepadSelectedList")
  );
  if (!ids.size) {
    toast("Select at least one deliverable to remove.");
    return;
  }

  deliverableNotepadSelectedEntryIds = deliverableNotepadSelectedEntryIds.filter(
    (id) => !ids.has(id)
  );
  renderDeliverableNotepadDialog();
}

async function exportSelectedDeliverablesToExcel() {
  const entryMap = new Map(deliverableNotepadEntries.map((item) => [item.id, item]));
  const selectedItems = deliverableNotepadSelectedEntryIds
    .map((id) => entryMap.get(id))
    .filter(Boolean);

  if (!selectedItems.length) {
    toast("Select at least one deliverable to export.");
    return;
  }

  if (!window.pywebview?.api?.export_deliverables_excel) {
    toast("Excel export is unavailable.");
    return;
  }

  const { entries, deliverableCount } =
    buildSelectedDeliverablesExcelRows(selectedItems);

  try {
    const response = await window.pywebview.api.export_deliverables_excel({
      entries,
    });
    if (response?.status === "success") {
      closeDlg("deliverableNotepadDlg");
      toast(
        `Exported ${deliverableCount} deliverable${
          deliverableCount === 1 ? "" : "s"
        } to Excel.`
      );
      return;
    }
    toast(response?.message || "Failed to export deliverables to Excel.");
  } catch (error) {
    console.error("Failed to export deliverables to Excel:", error);
    toast("Failed to export deliverables to Excel.");
  }
}

function openDeliverablesExcelDialog() {
  const dialog = document.getElementById("deliverableNotepadDlg");
  if (!dialog) return;

  deliverableNotepadEntries = buildDeliverableNotepadEntries(db);
  deliverableNotepadSelectedEntryIds = [];

  if (!deliverableNotepadEntries.length) {
    toast("No deliverables found on projects.");
    return;
  }

  renderDeliverableNotepadDialog();
  dialog.showModal();
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
  return getActiveAnchorDeliverable(project);
}

function getProjectListPriorityMeta(project) {
  const activeAnchorDeliverable = getActiveAnchorDeliverable(project);
  const activeIncompleteDeliverables = getProjectActiveDeliverables(project).filter(
    (deliverable) => !isFinished(deliverable)
  );
  if (!activeIncompleteDeliverables.length) {
    return {
      priorityDeliverable: activeAnchorDeliverable,
      hasIncompleteActiveWork: false,
      sortBucket: 2,
      sortDueDate: null,
      fallbackDueDate: parseDueStr(activeAnchorDeliverable?.due),
    };
  }

  const activeIncompleteWithDue = activeIncompleteDeliverables.filter((deliverable) =>
    parseDueStr(deliverable?.due)
  );
  if (activeIncompleteWithDue.length) {
    const priorityDeliverable = activeIncompleteWithDue.sort(compareDeliverablesByDue)[0];
    return {
      priorityDeliverable,
      hasIncompleteActiveWork: true,
      sortBucket: 0,
      sortDueDate: parseDueStr(priorityDeliverable?.due),
      fallbackDueDate: null,
    };
  }
  return {
    priorityDeliverable: activeIncompleteDeliverables[0],
    hasIncompleteActiveWork: true,
    sortBucket: 1,
    sortDueDate: null,
    fallbackDueDate: null,
  };
}

function getProjectListPriorityDeliverable(project) {
  return getProjectListPriorityMeta(project).priorityDeliverable;
}

function getOverviewDeliverables(project, { primaryId = "" } = {}) {
  const deliverables = getProjectDeliverables(project);
  if (!deliverables.length) return [];
  const out = deliverables.slice();
  const resolvedPrimaryId = String(primaryId || "").trim();
  const anchorId = resolvedPrimaryId || getActiveAnchorDeliverable(project)?.id || "";
  sortDeliverablesByPrimaryThenDueDesc(out, anchorId);
  return out;
}

function getProjectSortKey(project, projectListContext = null) {
  if (projectListContext && "anchorDueDate" in projectListContext) {
    return projectListContext.anchorDueDate;
  }
  const activeAnchorDeliverable =
    projectListContext?.activeAnchorDeliverable || getActiveAnchorDeliverable(project);
  return parseDueStr(activeAnchorDeliverable?.due);
}

function matchesProjectStatusFilter(deliverable, filter) {
  if (filter === "all") return true;
  if (filter === "incomplete") return !isFinished(deliverable);
  return hasStatus(deliverable, filter);
}

function matchesProjectDeliverablesFilter(
  deliverable,
  filter,
  activeAnchorDeliverable
) {
  if (filter === "all") return true;
  if (filter === "incomplete") return !isFinished(deliverable);
  if (filter === "active") {
    return deliverable?.active === true;
  }
  return true;
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
  const startOfLastWeek = new Date(startOfWeek);
  startOfLastWeek.setDate(startOfWeek.getDate() - 7);
  startOfLastWeek.setHours(0, 0, 0, 0);
  const endOfLastWeek = new Date(startOfWeek);
  endOfLastWeek.setMilliseconds(-1);

  if (filter === "lastWeek") return d >= startOfLastWeek && d <= endOfLastWeek;
  if (filter === "soon") return d >= startOfWeek && d <= endOfWeek;
  if (filter === "future") return d > endOfWeek;
  return true;
}

function getTimeframeFilterLabel(filter) {
  if (filter === "lastWeek") return "last week";
  if (filter === "soon") return "this week";
  if (filter === "future") return "upcoming weeks";
  return "";
}

function getLatestDeliverableDueDate(deliverables = []) {
  let latestDue = null;
  deliverables.forEach((deliverable) => {
    const due = parseDueStr(deliverable?.due);
    if (!due) return;
    if (!latestDue || due > latestDue) latestDue = due;
  });
  return latestDue;
}

function buildProjectTimeframeNote(
  filter,
  hasAdditionalFilters,
  activeAnchorMatchesTimeframe
) {
  const timeframeLabel = getTimeframeFilterLabel(filter);
  if (!timeframeLabel) return "";

  const prefix =
    hasAdditionalFilters
      ? `Showing deliverables based on ${timeframeLabel} and the active filters.`
      : `Showing deliverables based on ${timeframeLabel}.`;

  if (hasAdditionalFilters && activeAnchorMatchesTimeframe) {
    return `${prefix} Active deliverable does not match the current filters.`;
  }

  return `${prefix} Active deliverable is outside this timeframe.`;
}

function shouldSortCompletedProjectsLast() {
  return (
    separateDeliverableCompletionGroups &&
    currentSort.key === "due" &&
    dueFilter === "all" &&
    statusFilter === "all"
  );
}

function compareProjectListSortBuckets(a, b, projectListContextMap = null) {
  if (!shouldSortCompletedProjectsLast()) return 0;
  const aContext = projectListContextMap?.get(a) || null;
  const bContext = projectListContextMap?.get(b) || null;
  const aBucket = Number.isFinite(aContext?.sortBucket) ? aContext.sortBucket : 2;
  const bBucket = Number.isFinite(bContext?.sortBucket) ? bContext.sortBucket : 2;
  if (aBucket !== bBucket) return aBucket - bBucket;
  if (aBucket !== 2) return 0;

  const da = aContext?.fallbackDueDate || null;
  const dbb = bContext?.fallbackDueDate || null;
  if (!da && !dbb) return 0;
  if (!da) return 1;
  if (!dbb) return -1;
  return dbb - da;
}

function getProjectListRenderContext(project) {
  const projectListPriority = getProjectListPriorityMeta(project);
  const activeAnchorDeliverable = projectListPriority.priorityDeliverable;
  if (!activeAnchorDeliverable) return null;
  const overviewDeliverables = getOverviewDeliverables(project, {
    primaryId: activeAnchorDeliverable.id,
  });
  if (!overviewDeliverables.length) return null;

  const isTimeframeView = dueFilter !== "all";
  const timeframeDeliverables = isTimeframeView
    ? overviewDeliverables.filter((deliverable) =>
        matchesDueFilter(deliverable, dueFilter)
      )
    : overviewDeliverables;
  const statusMatchingDeliverables = timeframeDeliverables.filter(
    (deliverable) => matchesProjectStatusFilter(deliverable, statusFilter)
  );
  const filteredDeliverables = statusMatchingDeliverables.filter((deliverable) =>
    matchesProjectDeliverablesFilter(
      deliverable,
      deliverablesFilter,
      activeAnchorDeliverable
    )
  );
  const visibleDeliverables = isTimeframeView
    ? filteredDeliverables.slice().sort(compareDeliverablesByDueDesc)
    : filteredDeliverables;
  const matchesFilters = visibleDeliverables.length > 0;
  const activeAnchorMatchesTimeframe = isTimeframeView
    ? timeframeDeliverables.some(
        (deliverable) => deliverable.id === activeAnchorDeliverable.id
      )
    : true;
  const hasAdditionalFilters =
    statusFilter !== "all" || deliverablesFilter !== "all";
  const activeAnchorVisible = visibleDeliverables.some(
    (deliverable) => deliverable.id === activeAnchorDeliverable.id
  );
  const showTimeframeNote =
    isTimeframeView && visibleDeliverables.length > 0 && !activeAnchorVisible;

  return {
    activeAnchorDeliverable,
    hasIncompleteActiveWork: projectListPriority.hasIncompleteActiveWork,
    sortBucket: projectListPriority.sortBucket,
    sortDueDate: projectListPriority.sortDueDate,
    fallbackDueDate: projectListPriority.fallbackDueDate,
    timeframeDeliverables,
    statusMatchingDeliverables,
    visibleDeliverables,
    anchorDueDate: isTimeframeView
      ? getLatestDeliverableDueDate(visibleDeliverables)
      : projectListPriority.hasIncompleteActiveWork
        ? projectListPriority.sortDueDate
        : projectListPriority.fallbackDueDate,
    matchesFilters,
    timeframeLabel: getTimeframeFilterLabel(dueFilter),
    showTimeframeNote,
    timeframeNote: showTimeframeNote
      ? buildProjectTimeframeNote(
          dueFilter,
          hasAdditionalFilters,
          activeAnchorMatchesTimeframe
        )
      : "",
  };
}

function getProjectsFilterValue(filterKey) {
  if (filterKey === "timeframe") return dueFilter || "all";
  if (filterKey === "status") return statusFilter || "all";
  if (filterKey === "deliverables") return deliverablesFilter || "active";
  return "all";
}

function setProjectsFilterValue(filterKey, value) {
  if (filterKey === "timeframe") {
    dueFilter = value;
    if (currentSort.key === "due") {
      currentSort.dir = value === "all" ? "asc" : "desc";
    }
  } else if (filterKey === "status") {
    statusFilter = value;
  } else if (filterKey === "deliverables") {
    deliverablesFilter = value;
  }
}

function getProjectsFilterOptions(dropdown) {
  return dropdown
    ? Array.from(dropdown.querySelectorAll(".projects-filter-option"))
    : [];
}

function syncProjectsFilterDropdowns() {
  document.querySelectorAll(".projects-filter-dropdown").forEach((dropdown) => {
    const filterKey = dropdown.dataset.filterDropdown;
    const currentValue = getProjectsFilterValue(filterKey);
    const trigger = dropdown.querySelector(".projects-filter-trigger");
    const label = dropdown.querySelector(".projects-filter-trigger-label");
    const menu = dropdown.querySelector(".projects-filter-menu");
    const options = getProjectsFilterOptions(dropdown);
    const selectedOption =
      options.find((option) => option.dataset.filterValue === currentValue) ||
      options[0];
    const isOpen = dropdown.classList.contains("open");

    if (label && selectedOption) {
      label.textContent = selectedOption.textContent.trim();
    }
    if (trigger) {
      trigger.setAttribute("aria-expanded", String(isOpen));
    }
    if (menu) {
      menu.hidden = !isOpen;
    }

    options.forEach((option) => {
      const isSelected = option === selectedOption;
      option.classList.toggle("is-selected", isSelected);
      option.setAttribute("aria-checked", String(isSelected));
      option.tabIndex = isOpen ? 0 : -1;
    });
  });
}

function focusProjectsFilterSelectedOption(dropdown) {
  const selectedOption =
    dropdown?.querySelector('.projects-filter-option[aria-checked="true"]') ||
    dropdown?.querySelector(".projects-filter-option");
  selectedOption?.focus();
}

function setProjectsFilterDropdownState(
  dropdown,
  isOpen,
  { focusSelected = false } = {}
) {
  if (!dropdown) return;

  if (isOpen && openProjectsFilterDropdown && openProjectsFilterDropdown !== dropdown) {
    setProjectsFilterDropdownState(openProjectsFilterDropdown, false);
  }

  const menu = dropdown.querySelector(".projects-filter-menu");
  dropdown.classList.toggle("open", isOpen);
  menu?.classList.toggle("open", isOpen);

  if (isOpen) {
    openProjectsFilterDropdown = dropdown;
  } else if (openProjectsFilterDropdown === dropdown) {
    openProjectsFilterDropdown = null;
  }

  syncProjectsFilterDropdowns();

  if (isOpen && focusSelected) {
    focusProjectsFilterSelectedOption(dropdown);
  }
}

function moveProjectsFilterOptionFocus(dropdown, currentOption, direction) {
  const options = getProjectsFilterOptions(dropdown);
  if (!options.length) return;

  const currentIndex = options.indexOf(currentOption);
  const selectedIndex = options.findIndex(
    (option) => option.getAttribute("aria-checked") === "true"
  );
  const baseIndex =
    currentIndex >= 0 ? currentIndex : selectedIndex >= 0 ? selectedIndex : 0;
  const nextIndex = (baseIndex + direction + options.length) % options.length;

  options[nextIndex]?.focus();
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

function createDeliverableTaskCountBadge(deliverable) {
  const badge = el("div", {
    className: "deliverable-task-count-badge",
  });
  const icon = createIcon(CHECK_ICON_PATH, 12);
  icon.classList.add("deliverable-task-count-icon");
  const value = el("span", {
    className: "deliverable-task-count-value",
  });
  badge.append(icon, value);
  const stats = getTaskCompletionStats(deliverable);
  value.textContent = `${stats.completed}/${stats.total}`;
  return badge;
}

function updateDeliverableTaskStats(card, deliverable) {
  if (!card) return;

  const stats = getTaskCompletionStats(deliverable);
  const taskSummaryLabel = `${stats.completed} of ${stats.total} tasks complete`;

  const taskCountBadge = card.querySelector(".deliverable-task-count-badge");
  if (taskCountBadge) {
    const taskCountValue = taskCountBadge.querySelector(
      ".deliverable-task-count-value"
    );
    if (taskCountValue) {
      taskCountValue.textContent = `${stats.completed}/${stats.total}`;
    }
    taskCountBadge.classList.toggle("is-empty", stats.total === 0);
    taskCountBadge.classList.toggle(
      "is-complete",
      stats.total > 0 && stats.completed === stats.total
    );
    taskCountBadge.title = taskSummaryLabel;
    taskCountBadge.setAttribute("aria-label", taskSummaryLabel);
  }

  const progressSection = card.querySelector(".deliverable-progress-section");
  if (progressSection) {
    const barFill = progressSection.querySelector(".deliverable-progress-bar-fill");
    const progressText = progressSection.querySelector(".deliverable-progress-text");
    const percentageSpan = progressSection.querySelector(".percentage");
    const detailSpan = progressSection.querySelector(".detail");
    if (barFill) {
      barFill.style.width = `${stats.percentage}%`;
    }
    if (percentageSpan) {
      percentageSpan.textContent = `${stats.percentage}%`;
    }
    if (detailSpan) {
      detailSpan.textContent =
        stats.total > 0
          ? ` complete (${stats.completed}/${stats.total} tasks)`
          : " (no tasks)";
    }
    if (progressText) {
      progressText.className = `deliverable-progress-text ${
        stats.percentage >= 50 ? "high" : stats.percentage >= 25 ? "medium" : "low"
      }`;
    }
  }

  const tasksHeading = card.querySelector(".deliverable-tasks-preview-heading");
  if (tasksHeading) {
    tasksHeading.textContent =
      stats.total > 0 ? `Tasks (${stats.completed}/${stats.total}):` : "Tasks:";
  }
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

async function applyAttachmentOwnerEmailRefAtIndex(owner, slotIndex, nextRef, options = {}) {
  const {
    persistNow = false,
    modalCard = null,
    syncModalCard = false,
  } = options;
  const normalizedNext = normalizeEmailRef(nextRef);
  if (!owner || !normalizedNext) return false;
  if (!Number.isInteger(slotIndex)) return false;
  if (slotIndex < 0 || slotIndex >= MAX_DELIVERABLE_EMAIL_REFS) return false;

  const currentRefs = syncAttachmentOwnerEmailRefs(owner).slice();
  const currentLength = currentRefs.length;
  if (slotIndex > currentLength) return false;
  if (slotIndex === currentLength && currentLength >= MAX_DELIVERABLE_EMAIL_REFS) return false;
  if (slotIndex < currentLength && sameEmailRef(currentRefs[slotIndex], normalizedNext)) {
    return false;
  }

  const nextRefs = currentRefs.slice();
  if (slotIndex === currentLength) {
    nextRefs.push(normalizedNext);
  } else {
    nextRefs[slotIndex] = normalizedNext;
  }

  const normalizedNextRefs = normalizeEmailRefs(nextRefs);
  if (sameEmailRefList(currentRefs, normalizedNextRefs)) return false;
  await reconcileManagedEmailRefTransitions(currentRefs, normalizedNextRefs, { modalCard });

  owner.emailRefs = normalizedNextRefs;
  owner.emailRef = normalizedNextRefs[0] || null;
  if (syncModalCard && modalCard) setDeliverableCardEmailRefs(modalCard, normalizedNextRefs);
  if (persistNow) await save();
  return true;
}

async function applyDeliverableEmailRefAtIndex(deliverable, slotIndex, nextRef, options = {}) {
  return applyAttachmentOwnerEmailRefAtIndex(deliverable, slotIndex, nextRef, {
    ...options,
    syncModalCard: true,
  });
}

async function removeAttachmentOwnerEmailRefAtIndex(owner, slotIndex, options = {}) {
  const {
    persistNow = false,
    modalCard = null,
    syncModalCard = false,
  } = options;
  if (!owner || !Number.isInteger(slotIndex)) return false;
  if (slotIndex < 0 || slotIndex >= MAX_DELIVERABLE_EMAIL_REFS) return false;
  const currentRefs = syncAttachmentOwnerEmailRefs(owner).slice();
  if (slotIndex >= currentRefs.length) return false;

  const nextRefs = currentRefs.filter((_, index) => index !== slotIndex);
  if (sameEmailRefList(currentRefs, nextRefs)) return false;
  await reconcileManagedEmailRefTransitions(currentRefs, nextRefs, { modalCard });

  owner.emailRefs = nextRefs;
  owner.emailRef = nextRefs[0] || null;
  if (syncModalCard && modalCard) setDeliverableCardEmailRefs(modalCard, nextRefs);
  if (persistNow) await save();
  return true;
}

async function clearAttachmentOwnerEmailRefs(owner, options = {}) {
  const {
    persistNow = false,
    modalCard = null,
    syncModalCard = false,
  } = options;
  if (!owner) return false;
  const currentRefs = syncAttachmentOwnerEmailRefs(owner).slice();
  if (!currentRefs.length) return false;
  await reconcileManagedEmailRefTransitions(currentRefs, [], { modalCard });
  owner.emailRefs = [];
  owner.emailRef = null;
  if (syncModalCard && modalCard) setDeliverableCardEmailRefs(modalCard, []);
  if (persistNow) await save();
  return true;
}

async function removeDeliverableEmailRefAtIndex(deliverable, slotIndex, options = {}) {
  return removeAttachmentOwnerEmailRefAtIndex(deliverable, slotIndex, {
    ...options,
    syncModalCard: true,
  });
}

function renderAttachmentOwnerEmailSlots(container, descriptor = {}, options = {}) {
  if (!container || !descriptor?.owner) return;
  const {
    persistNow = false,
    modalCard = null,
    scope = "projects-tab",
    onChange = null,
  } = options;
  const { owner, project = null, kind = "deliverable" } = descriptor;

  syncAttachmentOwnerEmailRefs(owner);
  container.classList.add("deliverable-email-slots");
  container.innerHTML = "";
  const emailRefs = owner.emailRefs || [];
  const slotCount =
    emailRefs.length >= MAX_DELIVERABLE_EMAIL_REFS
      ? MAX_DELIVERABLE_EMAIL_REFS
      : Math.max(1, emailRefs.length + 1);

  const rerender = () =>
    renderAttachmentOwnerEmailSlots(container, descriptor, {
      persistNow,
      modalCard,
      scope,
      onChange,
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
      const changed = await applyAttachmentOwnerEmailRefAtIndex(
        owner,
        slotIndex,
        nextRef,
        {
          persistNow,
          modalCard,
          syncModalCard: kind === "deliverable",
        }
      );
      if (changed) {
        if (typeof onChange === "function") onChange();
        rerender();
      }
      return changed;
    };

    button.onclick = async (e) => {
      e.stopPropagation();
      if (button.classList.contains("is-busy")) return;
      const refs = syncAttachmentOwnerEmailRefs(owner);
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
        const context = buildAttachmentOwnerEmailContext(descriptor, scope);
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
          const refs = syncAttachmentOwnerEmailRefs(owner);
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
        const changed = await removeAttachmentOwnerEmailRefAtIndex(owner, slotIndex, {
          persistNow,
          modalCard,
          syncModalCard: kind === "deliverable",
        });
        if (changed) {
          if (typeof onChange === "function") onChange();
          rerender();
        }
      };
      slot.appendChild(removeBtn);
    }

    container.appendChild(slot);
  }
}

function renderDeliverableEmailSlots(container, deliverable, options = {}) {
  const { project = null, ...rest } = options;
  return renderAttachmentOwnerEmailSlots(
    container,
    {
      kind: "deliverable",
      owner: deliverable,
      deliverable,
      project,
    },
    rest
  );
}

let openDeliverableLinksContext = null;
let deliverableLinksPanelElements = null;
let deliverableLinksPanelPositionRaf = 0;
let deliverableLinksInputCommitPending = false;
let deliverableLinksPanelClosePending = false;
let deliverableLinksPanelActiveHost = null;
let deliverableLinksPanelActiveOwnerDocument = null;
let deliverableLinksPanelDetachOutsideListeners = null;
let deliverableLinksPanelFocusLossFrame = 0;

function updateDeliverableLinksTriggerState(trigger, links) {
  if (!trigger) return;
  const normalized = (Array.isArray(links) ? links : [])
    .map((link) => normalizeDeliverableLinkEntry(link))
    .filter(Boolean);
  const count = normalized.length;
  const preview = normalized
    .slice(0, 3)
    .map((link) => link.label || link.raw || link.url)
    .filter(Boolean)
    .join(", ");
  trigger.classList.toggle("is-linked", count > 0);
  trigger.textContent = "Links";
  trigger.setAttribute("aria-expanded", String(openDeliverableLinksContext?.trigger === trigger));
  trigger.setAttribute(
    "aria-label",
    count > 0 ? `Manage ${count} link${count === 1 ? "" : "s"}` : "Manage links"
  );
  trigger.title = count > 0 ? preview || "Manage links" : "Manage links";
}

function ensureDeliverableLinksPanel() {
  if (deliverableLinksPanelElements) return deliverableLinksPanelElements;

  const panel = el("div", {
    className: "deliverable-links-panel",
    hidden: true,
  });
  const list = el("div", { className: "deliverable-links-list" });
  const addRow = el("div", { className: "deliverable-links-add-row" });
  const addBullet = el("span", {
    className: "task-icon task-add-bullet deliverable-links-add-bullet",
    textContent: "○",
  });
  const fields = el("div", { className: "deliverable-links-add-fields" });
  addBullet.textContent = "+";
  const nameInput = el("input", {
    className: "deliverable-links-add-input deliverable-links-name-input",
    type: "text",
    placeholder: "Name (optional)",
    "aria-label": "Link name",
  });
  const linkInput = el("input", {
    className: "deliverable-links-add-input deliverable-links-link-input",
    type: "text",
    placeholder: "Paste a file path or URL...",
    "aria-label": "Link path or URL",
  });
  const scopeControl = el("label", {
    className: "deliverable-links-scope-control",
  });
  const scopeCheckbox = el("input", {
    className: "deliverable-links-scope-checkbox",
    type: "checkbox",
    "aria-label": "Save new link as project-wide",
  });
  const scopeText = el("span", {
    className: "deliverable-links-scope-label",
    textContent: "Project-wide",
  });

  scopeControl.append(scopeCheckbox, scopeText);
  fields.append(nameInput, linkInput, scopeControl);
  addRow.append(addBullet, fields);
  panel.append(list, addRow);
  document.body.appendChild(panel);

  deliverableLinksPanelElements = {
    panel,
    list,
    nameInput,
    linkInput,
    scopeControl,
    scopeCheckbox,
  };

  const handleInputKeydown = async (e) => {
    if (e.key !== "Enter") return;
    e.preventDefault();
    e.stopPropagation();
    await commitOpenDeliverableLinksInput({ refocus: true });
  };

  const handleInputClick = (e) => {
    e.stopPropagation();
  };

  [nameInput, linkInput].forEach((input) => {
    input.addEventListener("keydown", handleInputKeydown);
    input.addEventListener("click", handleInputClick);
  });
  scopeCheckbox.addEventListener("click", handleInputClick);
  scopeCheckbox.addEventListener("change", () => {
    if (!openDeliverableLinksContext) return;
    openDeliverableLinksContext.addScope =
      scopeCheckbox.checked ? "project" : "deliverable";
  });

  return deliverableLinksPanelElements;
}

function getPendingDeliverableLinkEntry() {
  const elements = ensureDeliverableLinksPanel();
  const raw = String(elements.linkInput.value || "")
    .split(/\r?\n/)
    .map((value) => value.trim())
    .find(Boolean);
  if (!raw) return null;
  return normalizeDeliverableLinkEntry({
    label: String(elements.nameInput.value || "").trim(),
    raw,
  });
}

function clearPendingDeliverableLinkEntry() {
  const elements = ensureDeliverableLinksPanel();
  elements.nameInput.value = "";
  elements.linkInput.value = "";
}

function focusDeliverableLinksPrimaryInput({ select = true } = {}) {
  const elements = ensureDeliverableLinksPanel();
  elements.nameInput.focus();
  if (select) elements.nameInput.select();
}

function syncDeliverableLinksAddScopeControl() {
  const elements = ensureDeliverableLinksPanel();
  const allowProjectScope = openDeliverableLinksContext?.allowProjectScope !== false;
  elements.scopeControl.hidden = !allowProjectScope;
  elements.scopeCheckbox.disabled = !allowProjectScope;
  elements.scopeCheckbox.checked =
    allowProjectScope &&
    (openDeliverableLinksContext?.addScope || "deliverable") === "project";
}

function getOpenDeliverableLinksPanelEntries() {
  if (!openDeliverableLinksContext) return [];
  return [
    ...openDeliverableLinksContext.getDeliverableLinks().map((link, index) => ({
      ...link,
      scope: "deliverable",
      index,
    })),
    ...openDeliverableLinksContext.getProjectLinks().map((link, index) => ({
      ...link,
      scope: "project",
      index,
    })),
  ];
}

async function updateOpenDeliverableLinksScopes({
  deliverableLinks = openDeliverableLinksContext?.getDeliverableLinks() || [],
  projectLinks = openDeliverableLinksContext?.getProjectLinks() || [],
} = {}) {
  if (!openDeliverableLinksContext?.updateScopedLinks) return null;
  return openDeliverableLinksContext.updateScopedLinks({
    deliverableLinks,
    projectLinks,
  });
}

async function moveOpenDeliverableLinkEntryScope(entry, nextScope) {
  if (!entry || !openDeliverableLinksContext) return false;
  const currentScope = entry.scope === "project" ? "project" : "deliverable";
  const targetScope = nextScope === "project" ? "project" : "deliverable";
  if (currentScope === targetScope) return false;

  const deliverableLinks = [...openDeliverableLinksContext.getDeliverableLinks()];
  const projectLinks = [...openDeliverableLinksContext.getProjectLinks()];
  const sourceLinks = currentScope === "project" ? projectLinks : deliverableLinks;
  const targetLinks = targetScope === "project" ? projectLinks : deliverableLinks;
  const [moved] = sourceLinks.splice(entry.index, 1);
  if (!moved) return false;
  targetLinks.push(moved);

  await updateOpenDeliverableLinksScopes({
    deliverableLinks,
    projectLinks,
  });
  renderOpenDeliverableLinksPanel();
  return true;
}

async function removeOpenDeliverableLinkEntry(entry) {
  if (!entry || !openDeliverableLinksContext) return false;
  const deliverableLinks = [...openDeliverableLinksContext.getDeliverableLinks()];
  const projectLinks = [...openDeliverableLinksContext.getProjectLinks()];
  const sourceLinks = entry.scope === "project" ? projectLinks : deliverableLinks;
  if (entry.index < 0 || entry.index >= sourceLinks.length) return false;
  sourceLinks.splice(entry.index, 1);

  await updateOpenDeliverableLinksScopes({
    deliverableLinks,
    projectLinks,
  });
  renderOpenDeliverableLinksPanel();
  return true;
}

function getDeliverableLinksPanelHost(trigger) {
  return trigger?.closest("dialog[open]") || document.body;
}

function getDeliverableLinksPanelOutsideTargets(trigger) {
  const ownerDocument = trigger?.ownerDocument || document;
  const host = getDeliverableLinksPanelHost(trigger);
  return [...new Set([host, ownerDocument].filter(Boolean))];
}

function isDeliverableLinksInteractionTarget(target) {
  if (!(target instanceof Node)) return false;
  const elements = deliverableLinksPanelElements;
  if (elements?.panel?.contains(target)) return true;
  return !!openDeliverableLinksContext?.trigger?.contains(target);
}

function detachDeliverableLinksOutsideListeners() {
  if (typeof deliverableLinksPanelDetachOutsideListeners === "function") {
    deliverableLinksPanelDetachOutsideListeners();
  }
  if (deliverableLinksPanelFocusLossFrame) {
    cancelAnimationFrame(deliverableLinksPanelFocusLossFrame);
    deliverableLinksPanelFocusLossFrame = 0;
  }
  deliverableLinksPanelDetachOutsideListeners = null;
  deliverableLinksPanelActiveHost = null;
  deliverableLinksPanelActiveOwnerDocument = null;
}

function scheduleDeliverableLinksPanelFocusLossCheck() {
  if (deliverableLinksPanelFocusLossFrame) {
    cancelAnimationFrame(deliverableLinksPanelFocusLossFrame);
  }

  const triggerSnapshot = openDeliverableLinksContext?.trigger || null;
  const ownerDocument =
    deliverableLinksPanelActiveOwnerDocument || triggerSnapshot?.ownerDocument || document;

  deliverableLinksPanelFocusLossFrame = requestAnimationFrame(() => {
    deliverableLinksPanelFocusLossFrame = 0;
    if (!openDeliverableLinksContext || openDeliverableLinksContext.trigger !== triggerSnapshot) {
      return;
    }
    const activeElement = ownerDocument?.activeElement || document.activeElement;
    if (isDeliverableLinksInteractionTarget(activeElement)) return;
    void requestDeliverableLinksPanelClose();
  });
}

function attachDeliverableLinksOutsideListeners(trigger) {
  detachDeliverableLinksOutsideListeners();

  const ownerDocument = trigger?.ownerDocument || document;
  const host = getDeliverableLinksPanelHost(trigger);
  const outsideTargets = getDeliverableLinksPanelOutsideTargets(trigger);
  deliverableLinksPanelActiveHost = host;
  deliverableLinksPanelActiveOwnerDocument = ownerDocument;

  const handleOutsidePointerLike = (e) => {
    if (!openDeliverableLinksContext) return;
    if (
      deliverableLinksPanelActiveHost instanceof HTMLDialogElement &&
      e.target === deliverableLinksPanelActiveHost
    ) {
      void requestDeliverableLinksPanelClose();
      return;
    }
    if (isDeliverableLinksInteractionTarget(e.target)) return;
    void requestDeliverableLinksPanelClose();
  };

  const handleOutsideFocusIn = (e) => {
    if (!openDeliverableLinksContext) return;
    if (isDeliverableLinksInteractionTarget(e.target)) return;
    void requestDeliverableLinksPanelClose();
  };

  const handleEscapeKey = (e) => {
    if (e.key !== "Escape" || !openDeliverableLinksContext) return;
    e.preventDefault();
    void requestDeliverableLinksPanelClose({ focusTrigger: true });
  };

  const handleFocusOut = () => {
    if (!openDeliverableLinksContext) return;
    scheduleDeliverableLinksPanelFocusLossCheck();
  };

  const handleViewportScroll = (e) => {
    if (!openDeliverableLinksContext) return;
    if (isDeliverableLinksInteractionTarget(e.target)) return;
    void requestDeliverableLinksPanelClose();
  };

  const handleViewportResize = () => {
    if (!openDeliverableLinksContext) return;
    scheduleDeliverableLinksPanelPosition();
  };

  outsideTargets.forEach((target) => {
    target.addEventListener("pointerdown", handleOutsidePointerLike, true);
    target.addEventListener("mousedown", handleOutsidePointerLike, true);
  });
  ensureDeliverableLinksPanel().panel.addEventListener("focusout", handleFocusOut);
  trigger.addEventListener("focusout", handleFocusOut);
  ownerDocument.addEventListener("focusin", handleOutsideFocusIn, true);
  ownerDocument.addEventListener("scroll", handleViewportScroll, true);
  ownerDocument.addEventListener("keydown", handleEscapeKey);
  ownerDocument.defaultView?.addEventListener("resize", handleViewportResize);

  deliverableLinksPanelDetachOutsideListeners = () => {
    outsideTargets.forEach((target) => {
      target.removeEventListener("pointerdown", handleOutsidePointerLike, true);
      target.removeEventListener("mousedown", handleOutsidePointerLike, true);
    });
    ensureDeliverableLinksPanel().panel.removeEventListener("focusout", handleFocusOut);
    trigger.removeEventListener("focusout", handleFocusOut);
    ownerDocument.removeEventListener("focusin", handleOutsideFocusIn, true);
    ownerDocument.removeEventListener("scroll", handleViewportScroll, true);
    ownerDocument.removeEventListener("keydown", handleEscapeKey);
    ownerDocument.defaultView?.removeEventListener("resize", handleViewportResize);
  };
}

function scheduleDeliverableLinksPanelPosition() {
  if (!openDeliverableLinksContext || deliverableLinksPanelPositionRaf) return;
  deliverableLinksPanelPositionRaf = requestAnimationFrame(() => {
    deliverableLinksPanelPositionRaf = 0;
    positionOpenDeliverableLinksPanel();
  });
}

function positionOpenDeliverableLinksPanel() {
  const elements = ensureDeliverableLinksPanel();
  if (!openDeliverableLinksContext) {
    elements.panel.hidden = true;
    return;
  }

  const trigger = openDeliverableLinksContext.trigger;
  if (!trigger?.isConnected) {
    void requestDeliverableLinksPanelClose();
    return;
  }

  const gap = 8;
  const rect = trigger.getBoundingClientRect();
  const panel = elements.panel;
  panel.hidden = false;
  panel.style.top = "0px";
  panel.style.left = "0px";
  panel.style.visibility = "hidden";

  const panelRect = panel.getBoundingClientRect();
  let top = rect.bottom + gap;
  if (top + panelRect.height > window.innerHeight - gap) {
    top = Math.max(gap, rect.top - panelRect.height - gap);
  }

  let left = rect.right - panelRect.width;
  left = Math.max(gap, Math.min(left, window.innerWidth - panelRect.width - gap));

  if (top + panelRect.height > window.innerHeight - gap) {
    top = Math.max(gap, window.innerHeight - panelRect.height - gap);
  }

  panel.style.top = `${Math.round(top)}px`;
  panel.style.left = `${Math.round(left)}px`;
  panel.style.visibility = "";
}

async function closeDeliverableLinksPanel({ focusTrigger = false } = {}) {
  if (deliverableLinksPanelClosePending) return false;
  const elements = ensureDeliverableLinksPanel();
  const trigger = openDeliverableLinksContext?.trigger || null;
  if (!trigger && !openDeliverableLinksContext) {
    detachDeliverableLinksOutsideListeners();
    elements.panel.hidden = true;
    elements.panel.style.visibility = "";
    return false;
  }

  deliverableLinksPanelClosePending = true;
  try {
    if (getPendingDeliverableLinkEntry()) {
      await commitOpenDeliverableLinksInput();
    }

    if (trigger) {
      trigger.classList.remove("open");
      trigger.setAttribute("aria-expanded", "false");
    }
    openDeliverableLinksContext = null;
    detachDeliverableLinksOutsideListeners();
    deliverableLinksInputCommitPending = false;
    clearPendingDeliverableLinkEntry();
    elements.panel.hidden = true;
    elements.panel.style.visibility = "";
    if (focusTrigger) {
      trigger?.focus();
    }
    return true;
  } finally {
    deliverableLinksPanelClosePending = false;
  }
}

async function requestDeliverableLinksPanelClose(options = {}) {
  return closeDeliverableLinksPanel(options);
}

function renderOpenDeliverableLinksPanel() {
  const elements = ensureDeliverableLinksPanel();
  if (!openDeliverableLinksContext) {
    elements.list.replaceChildren();
    elements.panel.hidden = true;
    return;
  }

  const { list } = elements;
  syncDeliverableLinksAddScopeControl();
  const allowProjectScope = openDeliverableLinksContext?.allowProjectScope !== false;
  const links = getOpenDeliverableLinksPanelEntries();
  list.replaceChildren();

  if (!links.length) {
    list.appendChild(
      el("div", {
        className: "deliverable-links-empty",
        textContent:
          allowProjectScope
            ? "No links yet. Add a name, choose whether it is project-wide, and paste a file path or URL below."
            : "No links yet. Add a name and paste a file path or URL below.",
      })
    );
  }

  links.forEach((entry) => {
    const item = el("div", {
      className: "deliverable-link-item",
    });
    const visibleLabel =
      String(entry.label || entry.raw || entry.url || "Link").trim() || "Link";
    const visibleMeta = String(
      entry.raw && entry.label && entry.raw !== entry.label
        ? entry.raw
        : entry.url || entry.raw || ""
    ).trim();
    const openBtn = el("button", {
      className: "deliverable-link-open",
      type: "button",
      title: entry.raw || entry.url || visibleLabel,
      "aria-label": `Open ${visibleLabel}`,
    });
    const label = el("span", {
      className: "deliverable-link-label",
      textContent: visibleLabel,
    });
    const meta = el("span", {
      className: "deliverable-link-meta",
      textContent: visibleMeta,
    });
    const actions = el("div", { className: "deliverable-link-actions" });
    const scopeControl = el("label", {
      className: "deliverable-link-scope-control",
      title:
        entry.scope === "project"
          ? "Turn off to make this link deliverable-only"
          : "Turn on to make this link project-wide",
    });
    const scopeCheckbox = el("input", {
      className: "deliverable-link-scope-checkbox",
      type: "checkbox",
      checked: entry.scope === "project",
      "aria-label": `Set ${visibleLabel} as project-wide`,
    });
    const scopeLabel = el("span", {
      className: "deliverable-link-scope-label",
      textContent: "Project-wide",
    });
    const removeBtn = el("button", {
      className: "deliverable-link-remove",
      type: "button",
      title: "Remove link",
      "aria-label": `Remove ${visibleLabel}`,
      textContent: "×",
    });

    removeBtn.textContent = "Remove";
    openBtn.append(label);
    if (visibleMeta) openBtn.append(meta);
    if (allowProjectScope) {
      scopeControl.append(scopeCheckbox, scopeLabel);
      actions.append(scopeControl);
    }
    actions.append(removeBtn);

    openBtn.addEventListener("click", async (e) => {
      e.stopPropagation();
      await openDeliverableLinkEntry(entry);
    });

    scopeCheckbox.addEventListener("click", (e) => {
      e.stopPropagation();
    });
    scopeCheckbox.addEventListener("change", async (e) => {
      e.stopPropagation();
      await moveOpenDeliverableLinkEntryScope(
        entry,
        scopeCheckbox.checked ? "project" : "deliverable"
      );
    });

    removeBtn.addEventListener("click", async (e) => {
      e.stopPropagation();
      await removeOpenDeliverableLinkEntry(entry);
      focusDeliverableLinksPrimaryInput({ select: false });
    });

    item.append(openBtn, actions);
    list.appendChild(item);
  });

  elements.panel.hidden = false;
  scheduleDeliverableLinksPanelPosition();
}

async function commitOpenDeliverableLinksInput({ refocus = false } = {}) {
  if (!openDeliverableLinksContext || deliverableLinksInputCommitPending) return false;
  const nextEntry = getPendingDeliverableLinkEntry();
  if (!nextEntry) return false;

  deliverableLinksInputCommitPending = true;
  clearPendingDeliverableLinkEntry();
  try {
    const deliverableLinks = [...openDeliverableLinksContext.getDeliverableLinks()];
    const projectLinks = [...openDeliverableLinksContext.getProjectLinks()];
    if ((openDeliverableLinksContext.addScope || "deliverable") === "project") {
      projectLinks.push(nextEntry);
    } else {
      deliverableLinks.push(nextEntry);
    }
    await updateOpenDeliverableLinksScopes({
      deliverableLinks,
      projectLinks,
    });
    renderOpenDeliverableLinksPanel();

    if (refocus) {
      setTimeout(() => {
        if (!openDeliverableLinksContext) return;
        focusDeliverableLinksPrimaryInput();
      }, 0);
    }
    return true;
  } finally {
    deliverableLinksInputCommitPending = false;
  }
}

async function openDeliverableLinkEntry(linkEntry) {
  const normalized = normalizeDeliverableLinkEntry(linkEntry);
  if (!normalized) {
    toast("No link is set.");
    return false;
  }

  const localPath = fileUrlToLocalPath(normalized.url || "");
  if (localPath) {
    if (!window.pywebview?.api?.reveal_path) {
      toast("Reveal path is unavailable.");
      return false;
    }
    try {
      const response = await window.pywebview.api.reveal_path(localPath);
      if (response?.status !== "success") {
        throw new Error(response?.message || "Unable to open linked path.");
      }
      toast("Opening path...");
      return true;
    } catch (e) {
      toast("Couldn't open linked path.");
      return false;
    }
  }

  const url = normalized.url || normalized.raw;
  if (!url) {
    toast("No link is set.");
    return false;
  }

  try {
    openExternalUrl(url);
    toast("Opening link...");
    return true;
  } catch (e) {
    toast("Couldn't open link.");
    return false;
  }
}

function openDeliverableLinksPanel(context) {
  const elements = ensureDeliverableLinksPanel();
  if (openDeliverableLinksContext?.trigger === context.trigger) {
    void requestDeliverableLinksPanelClose();
    return;
  }

  if (openDeliverableLinksContext?.trigger) {
    openDeliverableLinksContext.trigger.classList.remove("open");
    openDeliverableLinksContext.trigger.setAttribute("aria-expanded", "false");
  }

  openDeliverableLinksContext = context;
  openDeliverableLinksContext.addScope = "deliverable";
  const host = getDeliverableLinksPanelHost(context.trigger);
  if (elements.panel.parentElement !== host) {
    host.appendChild(elements.panel);
  }
  attachDeliverableLinksOutsideListeners(context.trigger);
  context.trigger.classList.add("open");
  context.trigger.setAttribute("aria-expanded", "true");
  clearPendingDeliverableLinkEntry();
  renderOpenDeliverableLinksPanel();
  positionOpenDeliverableLinksPanel();

  setTimeout(() => {
    if (openDeliverableLinksContext?.trigger !== context.trigger) return;
    focusDeliverableLinksPrimaryInput();
  }, 0);
}

function createAttachmentLinksControl(descriptor = {}, options = {}) {
  const {
    owner = null,
    kind = "deliverable",
    deliverable = kind === "deliverable" ? owner : descriptor.deliverable || null,
    project = null,
  } = descriptor;
  const {
    persistNow = false,
    modalCard = null,
    allowProjectScope = kind === "deliverable",
    onChange = null,
  } = options;
  const control = el("div", { className: "deliverable-link-control" });
  const trigger = el("button", {
    className: "deliverable-link-trigger",
    type: "button",
    textContent: "Links",
    "aria-expanded": "false",
    "aria-haspopup": "dialog",
  });

  const getOwnerLinks = () => {
    const normalized =
      kind === "deliverable" && modalCard
        ? getDeliverableCardLinks(modalCard)
        : normalizeDeliverableLinks(owner?.links, kind === "deliverable" ? owner?.linkPath || "" : "");
    if (owner) owner.links = normalized;
    if (kind === "deliverable" && modalCard) setDeliverableCardLinks(modalCard, normalized);
    return normalized;
  };

  const getProjectLinks = () => {
    if (!allowProjectScope) return [];
    const normalized = modalCard
      ? getModalProjectLinks()
      : normalizeDeliverableLinks(project?.links);
    if (project) project.links = normalized;
    if (modalCard) setModalProjectLinks(normalized);
    return normalized;
  };

  const updateScopedLinks = async ({
    deliverableLinks = getOwnerLinks(),
    projectLinks = getProjectLinks(),
  } = {}) => {
    const normalizedOwnerLinks = normalizeDeliverableLinks(
      deliverableLinks,
      kind === "deliverable" ? owner?.linkPath || "" : ""
    );
    const normalizedProjectLinks = allowProjectScope
      ? normalizeDeliverableLinks(projectLinks)
      : [];
    if (owner) owner.links = normalizedOwnerLinks;
    if (project && allowProjectScope) project.links = normalizedProjectLinks;
    if (kind === "deliverable" && modalCard) {
      setDeliverableCardLinks(modalCard, normalizedOwnerLinks);
    }
    if (allowProjectScope && modalCard) {
      setModalProjectLinks(normalizedProjectLinks);
    }
    if (persistNow) {
      await save();
    }
    if (typeof onChange === "function") onChange();
    syncState();
    return {
      deliverableLinks: normalizedOwnerLinks,
      projectLinks: normalizedProjectLinks,
    };
  };

  const setOwnerLinks = async (nextLinks) => {
    return (
      await updateScopedLinks({
        deliverableLinks: nextLinks,
      })
    ).deliverableLinks;
  };

  const setProjectLinks = async (nextLinks) => {
    return (
      await updateScopedLinks({
        projectLinks: nextLinks,
      })
    ).projectLinks;
  };

  const getCombinedLinks = () => {
    return [...getOwnerLinks(), ...getProjectLinks()];
  };

  const syncState = () => {
    updateDeliverableLinksTriggerState(trigger, getCombinedLinks());
    if (openDeliverableLinksContext?.trigger === trigger) {
      renderOpenDeliverableLinksPanel();
    }
  };

  trigger.addEventListener("click", (e) => {
    e.stopPropagation();
    document
      .querySelectorAll(".deliverable-status-dropdown.open")
      .forEach((openDropdown) => {
        setDeliverableStatusDropdownState(openDropdown, false);
      });
    openDeliverableLinksPanel({
      deliverable,
      project,
      modalCard,
      persistNow,
      trigger,
      allowProjectScope,
      ownerKind: kind,
      ownerLabel: getAttachmentOwnerLabel({
        kind,
        owner,
        deliverable,
      }),
      getDeliverableLinks: getOwnerLinks,
      setDeliverableLinks: setOwnerLinks,
      getProjectLinks,
      setProjectLinks,
      updateScopedLinks,
      syncState,
    });
  });

  control.appendChild(trigger);
  syncState();
  return control;
}

function createDeliverableLinksControl(deliverable, options = {}) {
  const { project = null, ...rest } = options;
  return createAttachmentLinksControl(
    {
      kind: "deliverable",
      owner: deliverable,
      deliverable,
      project,
    },
    {
      ...rest,
      allowProjectScope: true,
    }
  );
}

function sameAttachmentList(a, b) {
  const left = normalizeAttachments(a);
  const right = normalizeAttachments(b);
  if (left.length !== right.length) return false;
  for (let i = 0; i < left.length; i++) {
    const leftKey = getAttachmentEntryKey(left[i]);
    const rightKey = getAttachmentEntryKey(right[i]);
    if (leftKey !== rightKey) return false;
    if (String(left[i].description || "").trim() !== String(right[i].description || "").trim()) {
      return false;
    }
  }
  return true;
}

function getAttachmentOwnerSessionToken(descriptor = {}) {
  if (descriptor.modalCard) return descriptor.modalCard;
  if (descriptor.scope === "edit-modal") return { modalSession: true };
  return null;
}

function getAttachmentOwnerAttachments(descriptor = {}) {
  const {
    kind = "deliverable",
    owner = null,
    modalCard = null,
    scope = "projects-tab",
  } = descriptor;
  if (kind === "deliverable" && modalCard) {
    return getDeliverableCardAttachments(modalCard);
  }
  if (kind === "project" && scope === "edit-modal") {
    return getModalProjectAttachments();
  }
  if (!owner || typeof owner !== "object") return [];
  if (kind === "project") {
    syncProjectAttachmentFields(owner);
    return normalizeAttachments(owner.attachments);
  }
  syncAttachmentOwnerCompatFields(owner, {
    includeEmailRefs: true,
    legacyLinks:
      kind === "deliverable"
        ? normalizeDeliverableLinks(owner.links, owner.linkPath || "")
        : owner.links || [],
    legacyEmailRefs: owner.emailRefs,
    legacyEmailRef: owner.emailRef,
  });
  return normalizeAttachments(owner.attachments);
}

async function setAttachmentOwnerAttachments(
  descriptor = {},
  nextAttachments,
  options = {}
) {
  const {
    persistNow = false,
    onChange = null,
  } = options;
  const normalized = normalizeAttachments(nextAttachments);
  const current = getAttachmentOwnerAttachments(descriptor);
  if (sameAttachmentList(current, normalized)) {
    return current;
  }

  const sessionToken = getAttachmentOwnerSessionToken(descriptor);
  await reconcileManagedEmailRefTransitions(
    buildLegacyEmailRefsFromAttachments(current),
    buildLegacyEmailRefsFromAttachments(normalized),
    { modalCard: sessionToken }
  );

  const { kind = "deliverable", owner = null, modalCard = null, scope = "projects-tab" } =
    descriptor;
  if (kind === "deliverable" && modalCard) {
    setDeliverableCardAttachments(modalCard, normalized);
  }
  if (kind === "project" && scope === "edit-modal") {
    setModalProjectAttachments(normalized);
  }
  if (owner && typeof owner === "object") {
    owner.attachments = normalizeAttachments(normalized);
    if (kind === "project") {
      syncProjectAttachmentFields(owner);
    } else if (kind === "deliverable") {
      syncDeliverableAttachmentFields(owner);
    } else {
      syncAttachmentOwnerCompatFields(owner, {
        includeEmailRefs: true,
        legacyLinks:
          kind === "deliverable"
            ? normalizeDeliverableLinks(owner.links, owner.linkPath || "")
            : owner.links || [],
        legacyEmailRefs: owner.emailRefs,
        legacyEmailRef: owner.emailRef,
      });
    }
  }

  if (persistNow) {
    await save();
  }
  if (typeof onChange === "function") onChange(normalized);
  return normalized;
}

function createAttachmentTypeBadge(type = "path") {
  return el("span", {
    className: `attachment-type-badge is-${type}`,
    textContent: getAttachmentTypeLabel(type),
  });
}

async function openAttachmentEntry(attachment) {
  const normalized = normalizeAttachmentEntry(attachment);
  if (!normalized) {
    toast("No attachment is set.");
    return false;
  }

  if (normalized.type === "email") {
    return openDeliverableEmailRef(normalized.emailRef);
  }
  if (normalized.type === "url") {
    openExternalUrl(normalized.target);
    return true;
  }
  if (!window.pywebview?.api?.open_path) {
    toast("Open path is unavailable.");
    return false;
  }
  try {
    const result = await window.pywebview.api.open_path(normalized.target);
    if (result?.status && result.status !== "success") {
      throw new Error(result.message || "Unable to open attachment.");
    }
    return true;
  } catch (error) {
    toast(error?.message || "Unable to open attachment.");
    return false;
  }
}

let openAttachmentPanelContext = null;
let attachmentPanelElements = null;
let attachmentPanelPositionRaf = 0;
let attachmentPanelClosePending = false;
let attachmentPanelDetachOutsideListeners = null;
let attachmentPanelFocusLossFrame = 0;
let attachmentPanelActiveHost = null;
let attachmentPanelActiveOwnerDocument = null;

function updateAttachmentTriggerState(trigger, attachments) {
  if (!trigger) return;
  const normalized = normalizeAttachments(attachments);
  const count = normalized.length;
  const preview = normalized
    .slice(0, 3)
    .map(
      (attachment) =>
        String(
          attachment.description ||
            getAttachmentDescriptionFallback(attachment) ||
            "Attachment"
        ).trim() || "Attachment"
    )
    .join(", ");
  trigger.classList.toggle("has-attachments", count > 0);
  trigger.classList.toggle("open", openAttachmentPanelContext?.trigger === trigger);
  trigger.setAttribute("aria-expanded", String(openAttachmentPanelContext?.trigger === trigger));
  trigger.setAttribute(
    "aria-label",
    count > 0
      ? `Manage ${count} attachment${count === 1 ? "" : "s"}`
      : "Manage attachments"
  );
  trigger.title = count > 0 ? preview || "Manage attachments" : "Manage attachments";
}

function ensureAttachmentPanel() {
  if (attachmentPanelElements) return attachmentPanelElements;

  const panel = el("div", {
    className: "attachment-panel",
    hidden: true,
  });
  const list = el("div", { className: "attachment-list" });
  const composer = el("div", { className: "attachment-composer" });
  const descriptionInput = el("input", {
    className: "attachment-description-input",
    type: "text",
    placeholder: "Description (optional)",
    "aria-label": "Attachment description",
  });
  const pathInput = el("input", {
    className: "attachment-path-input",
    type: "text",
    placeholder: "Enter a path or URL",
    "aria-label": "Attachment path or URL",
  });
  const actionRow = el("div", { className: "attachment-action-row" });
  const chooseFileBtn = el("button", {
    className: "attachment-action-btn",
    type: "button",
    textContent: "Choose File",
  });
  const chooseFolderBtn = el("button", {
    className: "attachment-action-btn",
    type: "button",
    textContent: "Choose Folder",
  });
  const savePathBtn = el("button", {
    className: "attachment-action-btn",
    type: "button",
    textContent: "Save Path/URL",
  });
  const chooseEmailBtn = el("button", {
    className: "attachment-action-btn",
    type: "button",
    textContent: "Choose Email File",
  });

  actionRow.append(chooseFileBtn, chooseFolderBtn, savePathBtn, chooseEmailBtn);
  composer.append(descriptionInput, pathInput, actionRow);
  panel.append(list, composer);
  document.body.appendChild(panel);

  attachmentPanelElements = {
    panel,
    list,
    descriptionInput,
    pathInput,
    chooseFileBtn,
    chooseFolderBtn,
    savePathBtn,
    chooseEmailBtn,
  };

  const stopPropagation = (event) => {
    event.stopPropagation();
  };
  [descriptionInput, pathInput].forEach((input) => {
    input.addEventListener("click", stopPropagation);
  });

  return attachmentPanelElements;
}

function clearPendingAttachmentComposer() {
  const elements = ensureAttachmentPanel();
  elements.descriptionInput.value = "";
  elements.pathInput.value = "";
}

function focusAttachmentComposer() {
  const elements = ensureAttachmentPanel();
  elements.descriptionInput.focus();
  elements.descriptionInput.select();
}

async function addAttachmentToOpenPanel(entry) {
  if (!openAttachmentPanelContext) return false;
  const normalizedEntry = normalizeAttachmentEntry(entry);
  if (!normalizedEntry) return false;
  const nextAttachments = [
    ...openAttachmentPanelContext.getAttachments(),
    normalizedEntry,
  ];
  await openAttachmentPanelContext.setAttachments(nextAttachments);
  renderOpenAttachmentPanel();
  updateAttachmentTriggerState(
    openAttachmentPanelContext.trigger,
    openAttachmentPanelContext.getAttachments()
  );
  return true;
}

async function addDroppedEmailToAttachmentContext(context, event) {
  if (!context) return false;
  const resolved = await resolveEmailRefFromDrop(
    event,
    buildAttachmentOwnerEmailContext(context, context.scope || "projects-tab")
  );
  if (!resolved.emailRef) {
    showEmailLinkFallbackGuidance();
    return false;
  }
  const elements = ensureAttachmentPanel();
  const description =
    openAttachmentPanelContext?.trigger === context.trigger
      ? String(elements.descriptionInput.value || "").trim()
      : "";
  if (openAttachmentPanelContext?.trigger === context.trigger) {
    clearPendingAttachmentComposer();
  }
  await context.setAttachments([
    ...context.getAttachments(),
    {
      type: "email",
      description,
      emailRef: resolved.emailRef,
    },
  ]);
  if (openAttachmentPanelContext?.trigger === context.trigger) {
    renderOpenAttachmentPanel();
  }
  return true;
}

function buildPendingManualAttachment() {
  const elements = ensureAttachmentPanel();
  const target = String(elements.pathInput.value || "").trim();
  if (!target) return null;
  return normalizeAttachmentEntry({
    type: isWebAttachmentTarget(target) ? "url" : "path",
    target,
    description: String(elements.descriptionInput.value || "").trim(),
  });
}

async function chooseAttachmentFileFromPicker() {
  if (!window.pywebview?.api?.select_files) {
    toast("File picker is unavailable.");
    return null;
  }
  try {
    const result = await window.pywebview.api.select_files({
      allow_multiple: false,
    });
    if (result?.status !== "success" || !Array.isArray(result.paths) || !result.paths[0]) {
      return null;
    }
    const elements = ensureAttachmentPanel();
    return normalizeAttachmentEntry({
      type: "file",
      target: result.paths[0],
      description: String(elements.descriptionInput.value || "").trim(),
    });
  } catch {
    toast("Unable to open file picker.");
    return null;
  }
}

async function chooseAttachmentFolderFromPicker() {
  if (!window.pywebview?.api?.select_folder) {
    toast("Folder picker is unavailable.");
    return null;
  }
  try {
    const result = await window.pywebview.api.select_folder();
    if (result?.status !== "success" || !result.path) return null;
    const elements = ensureAttachmentPanel();
    return normalizeAttachmentEntry({
      type: "folder",
      target: result.path,
      description: String(elements.descriptionInput.value || "").trim(),
    });
  } catch {
    toast("Unable to open folder picker.");
    return null;
  }
}

async function addPendingAttachmentFromPathInput() {
  const attachment = buildPendingManualAttachment();
  if (!attachment) return false;
  clearPendingAttachmentComposer();
  return addAttachmentToOpenPanel(attachment);
}

function getAttachmentPanelHost(trigger) {
  return trigger?.closest("dialog[open]") || document.body;
}

function getAttachmentPanelOutsideTargets(trigger) {
  const ownerDocument = trigger?.ownerDocument || document;
  const host = getAttachmentPanelHost(trigger);
  return [...new Set([host, ownerDocument].filter(Boolean))];
}

function isAttachmentPanelInteractionTarget(target) {
  if (!(target instanceof Node)) return false;
  const elements = attachmentPanelElements;
  if (elements?.panel?.contains(target)) return true;
  return !!openAttachmentPanelContext?.trigger?.contains(target);
}

function detachAttachmentPanelOutsideListeners() {
  if (typeof attachmentPanelDetachOutsideListeners === "function") {
    attachmentPanelDetachOutsideListeners();
  }
  if (attachmentPanelFocusLossFrame) {
    cancelAnimationFrame(attachmentPanelFocusLossFrame);
    attachmentPanelFocusLossFrame = 0;
  }
  attachmentPanelDetachOutsideListeners = null;
  attachmentPanelActiveHost = null;
  attachmentPanelActiveOwnerDocument = null;
}

function scheduleAttachmentPanelFocusLossCheck() {
  if (attachmentPanelFocusLossFrame) {
    cancelAnimationFrame(attachmentPanelFocusLossFrame);
  }
  const triggerSnapshot = openAttachmentPanelContext?.trigger || null;
  const ownerDocument =
    attachmentPanelActiveOwnerDocument || triggerSnapshot?.ownerDocument || document;
  attachmentPanelFocusLossFrame = requestAnimationFrame(() => {
    attachmentPanelFocusLossFrame = 0;
    if (!openAttachmentPanelContext || openAttachmentPanelContext.trigger !== triggerSnapshot) {
      return;
    }
    const activeElement = ownerDocument?.activeElement || document.activeElement;
    if (isAttachmentPanelInteractionTarget(activeElement)) return;
    void requestAttachmentPanelClose();
  });
}

function attachAttachmentPanelOutsideListeners(trigger) {
  detachAttachmentPanelOutsideListeners();
  const ownerDocument = trigger?.ownerDocument || document;
  const host = getAttachmentPanelHost(trigger);
  const outsideTargets = getAttachmentPanelOutsideTargets(trigger);
  attachmentPanelActiveHost = host;
  attachmentPanelActiveOwnerDocument = ownerDocument;

  const handleOutsidePointerLike = (event) => {
    if (!openAttachmentPanelContext) return;
    if (
      attachmentPanelActiveHost instanceof HTMLDialogElement &&
      event.target === attachmentPanelActiveHost
    ) {
      void requestAttachmentPanelClose();
      return;
    }
    if (isAttachmentPanelInteractionTarget(event.target)) return;
    void requestAttachmentPanelClose();
  };
  const handleOutsideFocusIn = (event) => {
    if (!openAttachmentPanelContext) return;
    if (isAttachmentPanelInteractionTarget(event.target)) return;
    void requestAttachmentPanelClose();
  };
  const handleEscapeKey = (event) => {
    if (event.key !== "Escape" || !openAttachmentPanelContext) return;
    event.preventDefault();
    void requestAttachmentPanelClose({ focusTrigger: true });
  };
  const handleFocusOut = () => {
    if (!openAttachmentPanelContext) return;
    scheduleAttachmentPanelFocusLossCheck();
  };
  const handleViewportScroll = (event) => {
    if (!openAttachmentPanelContext) return;
    if (isAttachmentPanelInteractionTarget(event.target)) return;
    void requestAttachmentPanelClose();
  };
  const handleViewportResize = () => {
    if (!openAttachmentPanelContext) return;
    scheduleAttachmentPanelPosition();
  };

  outsideTargets.forEach((target) => {
    target.addEventListener("pointerdown", handleOutsidePointerLike, true);
    target.addEventListener("mousedown", handleOutsidePointerLike, true);
  });
  ensureAttachmentPanel().panel.addEventListener("focusout", handleFocusOut);
  trigger.addEventListener("focusout", handleFocusOut);
  ownerDocument.addEventListener("focusin", handleOutsideFocusIn, true);
  ownerDocument.addEventListener("scroll", handleViewportScroll, true);
  ownerDocument.addEventListener("keydown", handleEscapeKey);
  ownerDocument.defaultView?.addEventListener("resize", handleViewportResize);

  attachmentPanelDetachOutsideListeners = () => {
    outsideTargets.forEach((target) => {
      target.removeEventListener("pointerdown", handleOutsidePointerLike, true);
      target.removeEventListener("mousedown", handleOutsidePointerLike, true);
    });
    ensureAttachmentPanel().panel.removeEventListener("focusout", handleFocusOut);
    trigger.removeEventListener("focusout", handleFocusOut);
    ownerDocument.removeEventListener("focusin", handleOutsideFocusIn, true);
    ownerDocument.removeEventListener("scroll", handleViewportScroll, true);
    ownerDocument.removeEventListener("keydown", handleEscapeKey);
    ownerDocument.defaultView?.removeEventListener("resize", handleViewportResize);
  };
}

function scheduleAttachmentPanelPosition() {
  if (!openAttachmentPanelContext || attachmentPanelPositionRaf) return;
  attachmentPanelPositionRaf = requestAnimationFrame(() => {
    attachmentPanelPositionRaf = 0;
    positionOpenAttachmentPanel();
  });
}

function positionOpenAttachmentPanel() {
  const elements = ensureAttachmentPanel();
  if (!openAttachmentPanelContext) {
    elements.panel.hidden = true;
    return;
  }
  const trigger = openAttachmentPanelContext.trigger;
  if (!trigger?.isConnected) {
    void requestAttachmentPanelClose();
    return;
  }
  const gap = 8;
  const rect = trigger.getBoundingClientRect();
  const panel = elements.panel;
  panel.hidden = false;
  panel.style.top = "0px";
  panel.style.left = "0px";
  panel.style.visibility = "hidden";

  const panelRect = panel.getBoundingClientRect();
  let top = rect.bottom + gap;
  if (top + panelRect.height > window.innerHeight - gap) {
    top = Math.max(gap, rect.top - panelRect.height - gap);
  }
  let left = rect.right - panelRect.width;
  left = Math.max(gap, Math.min(left, window.innerWidth - panelRect.width - gap));
  if (top + panelRect.height > window.innerHeight - gap) {
    top = Math.max(gap, window.innerHeight - panelRect.height - gap);
  }
  panel.style.top = `${Math.round(top)}px`;
  panel.style.left = `${Math.round(left)}px`;
  panel.style.visibility = "";
}

async function closeAttachmentPanel({ focusTrigger = false } = {}) {
  if (attachmentPanelClosePending) return false;
  const elements = ensureAttachmentPanel();
  const trigger = openAttachmentPanelContext?.trigger || null;
  if (!trigger && !openAttachmentPanelContext) {
    detachAttachmentPanelOutsideListeners();
    elements.panel.hidden = true;
    elements.panel.style.visibility = "";
    return false;
  }
  attachmentPanelClosePending = true;
  try {
    clearPendingAttachmentComposer();
    if (trigger) {
      trigger.classList.remove("open");
      trigger.setAttribute("aria-expanded", "false");
    }
    openAttachmentPanelContext = null;
    detachAttachmentPanelOutsideListeners();
    elements.panel.hidden = true;
    elements.panel.style.visibility = "";
    if (focusTrigger) trigger?.focus();
    return true;
  } finally {
    attachmentPanelClosePending = false;
  }
}

async function requestAttachmentPanelClose(options = {}) {
  return closeAttachmentPanel(options);
}

function renderOpenAttachmentPanel() {
  const elements = ensureAttachmentPanel();
  if (!openAttachmentPanelContext) {
    elements.list.replaceChildren();
    elements.panel.hidden = true;
    return;
  }

  const attachments = openAttachmentPanelContext.getAttachments();
  elements.list.replaceChildren();
  if (!attachments.length) {
    elements.list.appendChild(
      el("div", {
        className: "attachment-empty",
        textContent: "No attachments yet. Add a file, folder, path, URL, or email below.",
      })
    );
  }

  attachments.forEach((attachment, index) => {
    const normalized = normalizeAttachmentEntry(attachment);
    if (!normalized) return;
    const row = el("div", { className: "attachment-item" });
    const openBtn = el("button", {
      className: "attachment-open",
      type: "button",
      title:
        normalized.type === "email"
          ? String(normalized.emailRef?.label || "Open email")
          : normalized.target || normalized.description || "Open attachment",
    });
    openBtn.append(
      createAttachmentTypeBadge(normalized.type),
      el("span", {
        className: "attachment-item-label",
        textContent:
          String(
            normalized.description ||
              getAttachmentDescriptionFallback(normalized) ||
              "Attachment"
          ).trim() || "Attachment",
      })
    );

    const removeBtn = el("button", {
      className: "attachment-remove",
      type: "button",
      textContent: "Remove",
      "aria-label": `Remove attachment ${index + 1}`,
    });

    openBtn.addEventListener("click", async (event) => {
      event.stopPropagation();
      await openAttachmentEntry(normalized);
    });
    removeBtn.addEventListener("click", async (event) => {
      event.stopPropagation();
      const nextAttachments = attachments.filter((_, entryIndex) => entryIndex !== index);
      await openAttachmentPanelContext.setAttachments(nextAttachments);
      renderOpenAttachmentPanel();
      updateAttachmentTriggerState(
        openAttachmentPanelContext.trigger,
        openAttachmentPanelContext.getAttachments()
      );
    });

    row.append(openBtn, removeBtn);
    elements.list.appendChild(row);
  });

  scheduleAttachmentPanelPosition();
}

function openAttachmentPanel(context) {
  const elements = ensureAttachmentPanel();
  if (openAttachmentPanelContext?.trigger === context.trigger) {
    void requestAttachmentPanelClose();
    return;
  }

  if (openAttachmentPanelContext?.trigger) {
    updateAttachmentTriggerState(
      openAttachmentPanelContext.trigger,
      openAttachmentPanelContext.getAttachments()
    );
  }

  openAttachmentPanelContext = context;
  const host = getAttachmentPanelHost(context.trigger);
  if (elements.panel.parentElement !== host) {
    host.appendChild(elements.panel);
  }
  attachAttachmentPanelOutsideListeners(context.trigger);
  clearPendingAttachmentComposer();
  renderOpenAttachmentPanel();
  positionOpenAttachmentPanel();
  updateAttachmentTriggerState(context.trigger, context.getAttachments());

  setTimeout(() => {
    if (openAttachmentPanelContext?.trigger !== context.trigger) return;
    focusAttachmentComposer();
  }, 0);
}

function createAttachmentControl(descriptor = {}, options = {}) {
  const {
    persistNow = false,
    onChange = null,
    className = "",
  } = options;
  const control = el("div", {
    className: `attachment-control ${className}`.trim(),
  });
  const trigger = el("button", {
    className: "attachment-trigger",
    type: "button",
    "aria-haspopup": "dialog",
    "aria-expanded": "false",
  });
  trigger.appendChild(createAttachmentTriggerIcon());

  const getAttachments = () => getAttachmentOwnerAttachments(descriptor);
  const setAttachments = async (nextAttachments) => {
    const normalized = await setAttachmentOwnerAttachments(descriptor, nextAttachments, {
      persistNow,
      onChange,
    });
    updateAttachmentTriggerState(trigger, normalized);
    return normalized;
  };
  const attachmentContext = {
    ...descriptor,
    trigger,
    getAttachments,
    setAttachments,
  };

  trigger.addEventListener("click", (event) => {
    event.stopPropagation();
    closeOpenDeliverableStatusDropdowns();
    closeOpenDeliverableToolDropdown();
    openAttachmentPanel(attachmentContext);
  });

  trigger.addEventListener("dragover", (event) => {
    event.preventDefault();
    event.stopPropagation();
    trigger.classList.add("is-dragover");
    if (event.dataTransfer) event.dataTransfer.dropEffect = "copy";
  });
  trigger.addEventListener("dragleave", (event) => {
    event.preventDefault();
    event.stopPropagation();
    trigger.classList.remove("is-dragover");
  });
  trigger.addEventListener("drop", async (event) => {
    event.preventDefault();
    event.stopPropagation();
    trigger.classList.remove("is-dragover");
    try {
      await addDroppedEmailToAttachmentContext(attachmentContext, event);
    } catch (error) {
      console.warn("Failed to process dropped email:", error);
      toast(error?.message || "Unable to process dropped email.");
    }
  });

  control.appendChild(trigger);
  updateAttachmentTriggerState(trigger, getAttachments());

  const elements = ensureAttachmentPanel();
  if (elements.chooseFileBtn.dataset.attachmentBound !== "true") {
    elements.chooseFileBtn.dataset.attachmentBound = "true";
    elements.chooseFileBtn.addEventListener("click", async (event) => {
      if (!openAttachmentPanelContext) return;
      event.stopPropagation();
      const entry = await chooseAttachmentFileFromPicker();
      if (!entry) return;
      clearPendingAttachmentComposer();
      await addAttachmentToOpenPanel(entry);
    });
  }
  if (elements.chooseFolderBtn.dataset.attachmentBound !== "true") {
    elements.chooseFolderBtn.dataset.attachmentBound = "true";
    elements.chooseFolderBtn.addEventListener("click", async (event) => {
      if (!openAttachmentPanelContext) return;
      event.stopPropagation();
      const entry = await chooseAttachmentFolderFromPicker();
      if (!entry) return;
      clearPendingAttachmentComposer();
      await addAttachmentToOpenPanel(entry);
    });
  }
  if (elements.chooseEmailBtn.dataset.attachmentBound !== "true") {
    elements.chooseEmailBtn.dataset.attachmentBound = "true";
    elements.chooseEmailBtn.addEventListener("click", async (event) => {
      if (!openAttachmentPanelContext) return;
      event.stopPropagation();
      const emailRef = await chooseEmailFileRefFromPicker();
      if (!emailRef) return;
      const description = String(ensureAttachmentPanel().descriptionInput.value || "").trim();
      clearPendingAttachmentComposer();
      await addAttachmentToOpenPanel({
        type: "email",
        description,
        emailRef,
      });
    });
  }
  if (elements.savePathBtn.dataset.attachmentBound !== "true") {
    elements.savePathBtn.dataset.attachmentBound = "true";
    elements.savePathBtn.addEventListener("click", async (event) => {
      if (!openAttachmentPanelContext) return;
      event.stopPropagation();
      await addPendingAttachmentFromPathInput();
    });
  }
  if (elements.pathInput.dataset.attachmentBound !== "true") {
    elements.pathInput.dataset.attachmentBound = "true";
    elements.pathInput.addEventListener("keydown", async (event) => {
      if (event.key !== "Enter") return;
      event.preventDefault();
      event.stopPropagation();
      await addPendingAttachmentFromPathInput();
    });
  }
  if (elements.panel.dataset.attachmentDropBound !== "true") {
    elements.panel.dataset.attachmentDropBound = "true";
    elements.panel.addEventListener("dragover", (event) => {
      event.preventDefault();
      if (event.dataTransfer) event.dataTransfer.dropEffect = "copy";
      elements.panel.classList.add("is-dragover");
    });
    elements.panel.addEventListener("dragleave", (event) => {
      event.preventDefault();
      elements.panel.classList.remove("is-dragover");
    });
    elements.panel.addEventListener("drop", async (event) => {
      event.preventDefault();
      elements.panel.classList.remove("is-dragover");
      try {
        await addDroppedEmailToAttachmentContext(openAttachmentPanelContext, event);
        renderOpenAttachmentPanel();
      } catch (error) {
        console.warn("Failed to process dropped email:", error);
        toast(error?.message || "Unable to process dropped email.");
      }
    });
  }

  return control;
}

let openDeliverableToolDropdown = null;
let deliverableToolDropdownGlobalHandlersBound = false;

function createDeliverableToolTriggerIcon() {
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("viewBox", "0 0 24 24");
  svg.setAttribute("fill", "none");
  svg.setAttribute("stroke", "currentColor");
  svg.setAttribute("stroke-width", "1.9");
  svg.setAttribute("stroke-linecap", "round");
  svg.setAttribute("stroke-linejoin", "round");
  svg.setAttribute("class", "deliverable-tool-trigger-icon");
  svg.innerHTML =
    '<path d="M21 7.5a5.5 5.5 0 0 1-7.6 5.08L7.5 18.5a2.12 2.12 0 1 1-3-3l5.92-5.9A5.5 5.5 0 0 1 16.5 3l-3 3 4.5 4.5 3-3Z"></path>';
  return svg;
}

function setDeliverableToolDropdownState(dropdown, isOpen) {
  if (!dropdown) return;

  const menu = dropdown.querySelector(".deliverable-tool-menu");
  const trigger = dropdown.querySelector(".deliverable-tool-trigger");
  const card = dropdown.closest(".deliverable-card-new");
  const cell = dropdown.closest(".cell-deliverables");
  const row = dropdown.closest(".row");

  menu?.classList.toggle("open", isOpen);
  trigger?.classList.toggle("open", isOpen);
  trigger?.setAttribute("aria-expanded", String(isOpen));
  dropdown.classList.toggle("open", isOpen);
  card?.classList.toggle("deliverable-menu-open", isOpen);
  cell?.classList.toggle("deliverable-menu-open", isOpen);
  row?.classList.toggle("deliverable-menu-open", isOpen);

  if (isOpen) {
    openDeliverableToolDropdown = dropdown;
  } else if (openDeliverableToolDropdown === dropdown) {
    openDeliverableToolDropdown = null;
  }
}

function closeOpenDeliverableToolDropdown({ except = null, focusTrigger = false } = {}) {
  if (!openDeliverableToolDropdown || openDeliverableToolDropdown === except) return;
  const dropdown = openDeliverableToolDropdown;
  setDeliverableToolDropdownState(dropdown, false);
  if (focusTrigger) {
    dropdown.querySelector(".deliverable-tool-trigger")?.focus();
  }
}

function closeOpenDeliverableStatusDropdowns(exceptDropdown = null) {
  document.querySelectorAll(".deliverable-status-dropdown.open").forEach((dropdown) => {
    if (dropdown !== exceptDropdown) {
      setDeliverableStatusDropdownState(dropdown, false);
    }
  });
}

function ensureDeliverableToolDropdownGlobalHandlers() {
  if (deliverableToolDropdownGlobalHandlersBound) return;
  deliverableToolDropdownGlobalHandlersBound = true;

  document.addEventListener("click", (e) => {
    if (!openDeliverableToolDropdown) return;
    if (openDeliverableToolDropdown.contains(e.target)) return;
    setDeliverableToolDropdownState(openDeliverableToolDropdown, false);
  });

  document.addEventListener("keydown", (e) => {
    if (e.key !== "Escape" || !openDeliverableToolDropdown) return;
    e.preventDefault();
    closeOpenDeliverableToolDropdown({ focusTrigger: true });
  });
}

function createDeliverableToolDropdown(deliverable, project, card) {
  ensureDeliverableToolDropdownGlobalHandlers();

  const dropdown = el("div", { className: "deliverable-tool-dropdown" });
  const trigger = el("button", {
    className: "deliverable-tool-trigger",
    type: "button",
    title: "Launch tools",
    "aria-label": "Launch tools",
    "aria-expanded": "false",
  });
  trigger.appendChild(createDeliverableToolTriggerIcon());

  const menu = el("div", { className: "deliverable-tool-menu" });
  getDeliverableToolMenuEntries().forEach((entry) => {
    const option = el("button", {
      className: "deliverable-tool-option",
      type: "button",
      textContent: entry.menuLabel || entry.label,
      "data-shared-tool-id": entry.id,
      "data-launch-type": entry.launchType,
    });
    option.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      setDeliverableToolDropdownState(dropdown, false);
      const launchContext = buildProjectsTabToolLaunchContext(project, deliverable);
      launchSharedToolCard(entry.id, launchContext);
    });
    menu.appendChild(option);
  });

  trigger.addEventListener("click", (e) => {
    e.preventDefault();
    e.stopPropagation();
    const isOpen = !dropdown.classList.contains("open");

    closeOpenDeliverableToolDropdown({ except: dropdown });
    closeOpenDeliverableStatusDropdowns();
    if (openAttachmentPanelContext) {
      void requestAttachmentPanelClose();
    }

    setDeliverableToolDropdownState(dropdown, isOpen);
  });

  dropdown.append(trigger, menu);
  card?.classList.remove("deliverable-menu-open");
  return dropdown;
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
  const taskCountBadge = createDeliverableTaskCountBadge(deliverable);
  leftSection.appendChild(taskCountBadge);

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

  // Right section: attachments + expand/contract controls
  const actions = el("div", { className: "deliverable-header-actions" });
  const expandToggle = createExpandToggle(card);
  const toolDropdown = createDeliverableToolDropdown(deliverable, project, card);
  const attachmentControl = createAttachmentControl(
    {
      kind: "deliverable",
      owner: deliverable,
      deliverable,
      project,
      scope: "projects-tab",
    },
    {
      persistNow: true,
      onChange: () => updateDeliverableWorkItemUi(card, deliverable),
    }
  );
  actions.append(attachmentControl, toolDropdown, expandToggle);
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
  renderDeliverableStatusBadges(container, deliverable);
  return container;
}

function renderDeliverableStatusBadges(container, deliverable) {
  if (!container) return;
  container.replaceChildren();
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
}

function setDeliverableStatusDropdownState(dropdown, isOpen) {
  if (!dropdown) return;

  const menu = dropdown.querySelector(".deliverable-status-menu");
  const trigger = dropdown.querySelector(".deliverable-status-trigger");
  const card = dropdown.closest(".deliverable-card-new");
  const cell = dropdown.closest(".cell-deliverables");
  const row = dropdown.closest(".row");

  menu?.classList.toggle("open", isOpen);
  trigger?.classList.toggle("open", isOpen);
  trigger?.setAttribute("aria-expanded", String(isOpen));
  dropdown.classList.toggle("open", isOpen);
  card?.classList.toggle("deliverable-menu-open", isOpen);
  cell?.classList.toggle("deliverable-menu-open", isOpen);
  row?.classList.toggle("deliverable-menu-open", isOpen);
}

function createStatusDropdown(deliverable, project, card) {
  const availableStatuses = ["Waiting", "Working", "Pending Review", "Complete", "Delivered"];
  const dropdown = el("div", { className: "deliverable-status-dropdown" });

  // Trigger button - compact ellipsis icon
  const trigger = el("button", {
    className: "deliverable-status-trigger",
    type: "button",
    title: "Status options",
    "aria-label": "Status options",
    "aria-expanded": "false"
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
      setDeliverableStatusDropdownState(dropdown, false);
      renderProjectsPreservingExpandedDeliverables();
    });

    const label = el("span", { textContent: status });
    option.append(checkbox, label);
    menu.appendChild(option);
  });

  // Toggle dropdown
  trigger.addEventListener("click", (e) => {
    e.stopPropagation();
    const isOpen = !dropdown.classList.contains("open");

    closeOpenDeliverableStatusDropdowns(dropdown);
    closeOpenDeliverableToolDropdown();
    if (openAttachmentPanelContext) {
      void requestAttachmentPanelClose();
    }

    setDeliverableStatusDropdownState(dropdown, isOpen);
  });

  // Close dropdown when clicking outside
  document.addEventListener("click", (e) => {
    if (!dropdown.contains(e.target)) {
      setDeliverableStatusDropdownState(dropdown, false);
    }
  });

  dropdown.append(trigger, menu);
  return dropdown;
}

function createDeliverableStatusSection(deliverable, project, card) {
  const statusSection = el("div", { className: "deliverable-status-row" });
  const statusBadges = createStatusBadges(deliverable);
  const statusInlineGroup = el("div", {
    className: "deliverable-status-inline-group",
  });
  const pinnedHost = el("div", {
    className: "deliverable-pinned-inline-group",
    hidden: true,
  });
  const statusDropdown = createStatusDropdown(deliverable, project, card);

  renderDeliverablePinnedPreview(pinnedHost, deliverable);
  statusInlineGroup.append(pinnedHost, statusDropdown);
  statusSection.append(statusBadges, statusInlineGroup);
  return statusSection;
}

function createNotesSectionLegacy(deliverable, project) {
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

function createTasksPreviewLegacy(deliverable, card) {
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

  // Helper to keep header badge, progress, and preview heading in sync.
  const updateStatsDisplay = () => {
    updateDeliverableTaskStats(card, deliverable);
  };

  const renderTaskList = () => {
    list.innerHTML = "";

    deliverable.tasks.forEach((task, index) => {
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

        deliverable.tasks.splice(index, 1);

        updateStatsDisplay();
        await save();
        renderTaskList();
      });

      item.append(iconSpan, textSpan, deleteBtn);

      // Make task clickable to toggle completion (but not on delete button)
      item.addEventListener("click", async (e) => {
        if (e.target === deleteBtn) return;
        e.stopPropagation();

        // Toggle done state
        taskObj.done = !taskObj.done;

        // Update the task in the array
        if (typeof deliverable.tasks[index] === "string") {
          deliverable.tasks[index] = { text: deliverable.tasks[index], done: taskObj.done };
        } else {
          deliverable.tasks[index].done = taskObj.done;
        }

        // Update UI
        item.classList.toggle("done", taskObj.done);
        item.classList.toggle("undone", !taskObj.done);
        iconSpan.textContent = taskObj.done ? "✓" : "○";

        updateStatsDisplay();
        await save();
      });

      list.appendChild(item);
    });

    // Add new task input at the bottom
    const addTaskRow = el("div", { className: "task-add-row" });
    const bulletSpan = el("span", { className: "task-icon task-add-bullet", textContent: "○" });
    const taskInput = el("input", {
      className: "task-add-input",
      type: "text",
      placeholder: "Add a task..."
    });

    let isCommittingTask = false;
    const commitPendingTask = async ({ refocus = false } = {}) => {
      const taskText = taskInput.value.trim();
      if (!taskText || isCommittingTask) return false;

      isCommittingTask = true;
      taskInput.value = "";
      try {
        deliverable.tasks.push({ text: taskText, done: false });

        updateStatsDisplay();
        await save();

        renderTaskList();

        if (refocus) {
          setTimeout(() => {
            const newTaskInput = list.querySelector(".task-add-input");
            if (newTaskInput) {
              newTaskInput.focus();
              newTaskInput.select();
            }
          }, 0);
        }
        return true;
      } finally {
        isCommittingTask = false;
      }
    };

    taskInput.addEventListener("keydown", async (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        e.stopPropagation();
        await commitPendingTask({ refocus: true });
      }
    });

    taskInput.addEventListener("blur", () => {
      if (!taskInput.value.trim()) return;
      setTimeout(() => {
        void commitPendingTask({ refocus: false });
      }, 0);
    });

    taskInput.addEventListener("click", (e) => {
      e.stopPropagation();
    });

    addTaskRow.append(bulletSpan, taskInput);
    list.appendChild(addTaskRow);
  };

  renderTaskList();

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

function getExpandedProjectDeliverableIds() {
  const tbody = document.getElementById("tbody");
  if (!tbody) return [];

  const expandedIds = [];
  tbody
    .querySelectorAll(".deliverable-card-new[data-deliverable-id]")
    .forEach((card) => {
      if (card.classList.contains("details-collapsed")) return;
      const deliverableId = String(card.dataset.deliverableId || "").trim();
      if (deliverableId) expandedIds.push(deliverableId);
    });
  return [...new Set(expandedIds)];
}

function restoreExpandedProjectDeliverables(expandedIds = []) {
  const tbody = document.getElementById("tbody");
  if (!tbody || !Array.isArray(expandedIds) || !expandedIds.length) return;

  const expandedSet = new Set(
    expandedIds.map((id) => String(id || "").trim()).filter(Boolean)
  );
  if (!expandedSet.size) return;

  tbody
    .querySelectorAll(".deliverable-card-new[data-deliverable-id]")
    .forEach((card) => {
      const deliverableId = String(card.dataset.deliverableId || "").trim();
      if (expandedSet.has(deliverableId)) {
        setDeliverableDetailsCollapsed(card, false);
      }
    });
}

function renderProjectsPreservingExpandedDeliverables() {
  const expandedIds = getExpandedProjectDeliverableIds();
  render();
  restoreExpandedProjectDeliverables(expandedIds);
}

function renderDeliverableCardLegacy(deliverable, isPrimary, project) {
  syncDeliverableWorkItemFields(deliverable);
  const deliverableId = String(deliverable?.id || createId("dlv")).trim();
  if (!deliverable?.id) deliverable.id = deliverableId;
  const card = el("div", {
    className: `deliverable-card-new ${isPrimary ? "is-primary" : ""} details-collapsed`
  });
  card.dataset.deliverableId = deliverableId;

  // Header: name + due badge + expand toggle (pass card for toggle)
  const header = createCardHeader(deliverable, isPrimary, card, project);

  // Progress: bar + percentage text
  const progress = createProgressSection(deliverable);

  // Status section: badges + grouped pinned preview / dropdown controls
  const statusSection = createDeliverableStatusSection(
    deliverable,
    project,
    card
  );

  // Tasks preview (2-3 tasks, now clickable)
  const tasksPreview = createTasksPreviewLegacy(deliverable, card);

  // Notes section (always visible when expanded, no toggle)
  const notesSection = createNotesSectionLegacy(deliverable, project);

  card.append(header, progress, statusSection, tasksPreview, notesSection);
  updateDeliverableTaskStats(card, deliverable);
  return card;
}

function renderDeliverablePinnedPreview(container, deliverable) {
  if (!container) return;
  const previewItems = getPinnedDeliverablePreviewItems(deliverable);
  container.replaceChildren();
  container.hidden = previewItems.length === 0;
  if (!previewItems.length) return;

  previewItems.forEach((previewItem) => {
    const pillIcon = createIcon(
      previewItem.type === "task" ? CHECK_ICON_PATH : NOTE_ICON_PATH,
      10
    );
    pillIcon.classList.add("deliverable-pinned-inline-pill-icon");
    pillIcon.setAttribute("aria-hidden", "true");
    const contentPill = el("span", {
      className: `deliverable-pinned-inline-pill ${
        previewItem.type === "task" ? "is-task" : "is-note"
      } ${previewItem.done ? "done" : ""}`,
      title: previewItem.text || "",
    });
    contentPill.append(
      pillIcon,
      el("span", {
        className: "deliverable-pinned-inline-pill-text",
        textContent: previewItem.text || "",
      })
    );
    container.appendChild(contentPill);
  });
}

function updateDeliverableWorkItemUi(card, deliverable) {
  updateDeliverableTaskStats(card, deliverable);
  renderDeliverableStatusBadges(
    card?.querySelector(".deliverable-status-badges"),
    deliverable
  );
  renderDeliverablePinnedPreview(
    card?.querySelector(".deliverable-pinned-inline-group"),
    deliverable
  );
}

function createWorkItemDoneCheckbox({
  checked = false,
  ariaLabel = "Mark task complete",
  onToggle = async () => {},
}) {
  const checkbox = el("input", {
    className: "work-item-checkbox",
    type: "checkbox",
    checked: !!checked,
    "aria-label": ariaLabel,
  });

  checkbox.addEventListener("click", (e) => {
    e.stopPropagation();
  });

  checkbox.addEventListener("change", async (e) => {
    e.stopPropagation();
    await onToggle(checkbox.checked);
  });

  return checkbox;
}

function createInlineWorkItemTextControl({
  textClassName = "",
  value = "",
  fallbackText = "",
  title = "Click to edit",
  ariaLabel = "Edit text",
  onCommit = async () => {},
}) {
  const wrapper = el("div", { className: "work-item-text-wrap" });
  const trigger = el("button", {
    className: `work-item-text-trigger ${textClassName}`.trim(),
    type: "button",
    title,
    "aria-label": ariaLabel,
  });

  let currentText = String(value || "");
  let input = null;
  let isFinishing = false;

  const syncTriggerText = (nextText = currentText) => {
    currentText = String(nextText || "");
    trigger.textContent = currentText || fallbackText;
  };

  const finishEditing = async (mode = "commit") => {
    if (!input || isFinishing) return;
    isFinishing = true;

    const activeInput = input;
    const previousText = String(activeInput.dataset.previousText || "");
    const trimmedText = activeInput.value.trim();
    input = null;

    activeInput.remove();
    wrapper.classList.remove("is-editing");
    trigger.hidden = false;

    if (mode === "cancel") {
      syncTriggerText(previousText);
      isFinishing = false;
      return;
    }

    const nextText = trimmedText || previousText;
    syncTriggerText(nextText);

    try {
      if (nextText !== previousText) {
        await onCommit(nextText, previousText);
      }
    } finally {
      isFinishing = false;
    }
  };

  trigger.addEventListener("click", (e) => {
    e.stopPropagation();

    if (input) {
      input.focus();
      input.select();
      return;
    }

    const previousText = String(currentText || "");
    input = el("input", {
      className: "work-item-text-input",
      type: "text",
      value: previousText,
    });
    input.dataset.previousText = previousText;

    wrapper.classList.add("is-editing");
    trigger.hidden = true;
    wrapper.appendChild(input);

    input.focus();
    input.select();

    input.addEventListener("click", (event) => {
      event.stopPropagation();
    });

    input.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        event.stopPropagation();
        void finishEditing("commit");
      } else if (event.key === "Escape") {
        event.preventDefault();
        event.stopPropagation();
        void finishEditing("cancel");
      }
    });

    input.addEventListener("blur", () => {
      void finishEditing("commit");
    });
  });

  syncTriggerText();
  wrapper.appendChild(trigger);
  return wrapper;
}

function createNotesSection(deliverable, card, project = null) {
  syncDeliverableNoteFields(deliverable);
  const section = el("div", { className: "deliverable-notes-section" });
  const header = el("div", { className: "deliverable-notes-header-simple" });
  const titleText = el("span", {
    className: "deliverable-notes-label",
    textContent: "Notes",
  });
  header.appendChild(titleText);

  const list = el("div", { className: "deliverable-notes-list-edit" });

  const renderNoteList = () => {
    list.innerHTML = "";

    getPinnedDeliverableNoteEntries(deliverable).forEach(({ item: noteItem, index }) => {
      deliverable.noteItems[index] = normalizeDeliverableNoteItem(
        deliverable.noteItems[index]
      );
      const liveNote = deliverable.noteItems[index];
      const row = el("div", { className: "deliverable-note-item" });
      const icon = el("span", {
        className: "note-icon",
        "aria-hidden": "true",
      });
      icon.appendChild(createIcon(NOTE_ICON_PATH, 12));
      const content = el("div", { className: "work-item-content" });
      const textControl = createInlineWorkItemTextControl({
        textClassName: "note-text",
        value: liveNote.text || "",
        fallbackText: "Note",
        title: "Click to edit note",
        ariaLabel: "Edit note text",
        onCommit: async (nextText) => {
          liveNote.text = nextText;
          syncDeliverableNoteFields(deliverable);
          updateDeliverableWorkItemUi(card, deliverable);
          renderNoteList();
          await save();
        },
      });
      const attachments = createWorkItemAttachmentControls({
        kind: "note",
        owner: liveNote,
        deliverable,
        project,
        persistNow: true,
        scope: "projects-tab",
        onChange: () => {
          updateDeliverableWorkItemUi(card, deliverable);
        },
      });
      const pinBtn = createWorkItemPinButton({
        pinned: !!liveNote.pinned,
        titlePinned: "Unpin note",
        titleUnpinned: "Pin note",
        onToggle: async (nextPinned) => {
          liveNote.pinned = nextPinned;
          if (
            deliverable.noteItems[index] &&
            typeof deliverable.noteItems[index] === "object" &&
            !Array.isArray(deliverable.noteItems[index])
          ) {
            deliverable.noteItems[index].pinned = nextPinned;
          }
          syncDeliverableWorkItemFields(deliverable);
          renderNoteList();
          updateDeliverableWorkItemUi(card, deliverable);
          await save();
          return nextPinned;
        },
      });
      const deleteBtn = el("button", {
        className: "task-delete-btn",
        type: "button",
        title: "Remove note",
        textContent: "×",
      });
      const actions = el("div", { className: "work-item-actions" });
      deleteBtn.addEventListener("click", async (e) => {
        e.stopPropagation();
        await clearAttachmentOwnerEmailRefs(liveNote);
        deliverable.noteItems.splice(index, 1);
        syncDeliverableNoteFields(deliverable);
        updateDeliverableWorkItemUi(card, deliverable);
        renderNoteList();
        await save();
      });
      actions.append(attachments, pinBtn, deleteBtn);
      content.append(textControl);
      row.append(icon, content, actions);
      list.appendChild(row);
    });

    const addNoteRow = el("div", { className: "task-add-row" });
    const noteBullet = el("span", {
      className: "note-icon note-add-bullet",
      "aria-hidden": "true",
    });
    noteBullet.appendChild(createIcon(NOTE_ICON_PATH, 12));
    const noteInput = el("input", {
      className: "task-add-input note-add-input",
      type: "text",
      placeholder: "Add a note...",
    });

    let isCommittingNote = false;
    const commitPendingNote = async ({ refocus = false } = {}) => {
      const noteText = noteInput.value.trim();
      if (!noteText || isCommittingNote) return false;

      isCommittingNote = true;
        noteInput.value = "";
      try {
        deliverable.noteItems.push({
          text: noteText,
          pinned: false,
          attachments: [],
        });
        syncDeliverableNoteFields(deliverable);
        updateDeliverableWorkItemUi(card, deliverable);
        await save();
        renderNoteList();

        if (refocus) {
          setTimeout(() => {
            const newNoteInput = list.querySelector(".note-add-input");
            if (newNoteInput) {
              newNoteInput.focus();
              newNoteInput.select();
            }
          }, 0);
        }
        return true;
      } finally {
        isCommittingNote = false;
      }
    };

    noteInput.addEventListener("keydown", async (e) => {
      if (e.key !== "Enter") return;
      e.preventDefault();
      e.stopPropagation();
      await commitPendingNote({ refocus: true });
    });

    noteInput.addEventListener("blur", () => {
      if (!noteInput.value.trim()) return;
      setTimeout(() => {
        void commitPendingNote({ refocus: false });
      }, 0);
    });

    noteInput.addEventListener("click", (e) => {
      e.stopPropagation();
    });

    addNoteRow.append(noteBullet, noteInput);
    list.appendChild(addNoteRow);
  };

  renderNoteList();
  section.append(header, list);
  return section;
}

function createTasksPreview(deliverable, card, project = null) {
  syncDeliverableWorkItemFields(deliverable);
  const container = el("div", { className: "deliverable-tasks-preview" });

  if (!deliverable.tasks) {
    deliverable.tasks = [];
  }

  const stats = getTaskCompletionStats(deliverable);
  const heading = el("div", {
    className: "deliverable-tasks-preview-heading",
    textContent:
      deliverable.tasks.length > 0
        ? `Tasks (${stats.completed}/${stats.total}):`
        : "Tasks:",
  });

  const list = el("div", { className: "deliverable-tasks-preview-list" });

  const updateStatsDisplay = () => {
    updateDeliverableWorkItemUi(card, deliverable);
  };

  const renderTaskList = () => {
    list.innerHTML = "";

    getPinnedDeliverableTaskEntries(deliverable).forEach(({ item: taskObj, index }) => {
      deliverable.tasks[index] = normalizeTask(deliverable.tasks[index]);
      const liveTask = deliverable.tasks[index];
      const item = el("div", {
        className: `deliverable-task-item ${liveTask.done ? "done" : "undone"}`,
      });
      const checkbox = createWorkItemDoneCheckbox({
        checked: !!liveTask.done,
        ariaLabel: liveTask.done ? "Mark task incomplete" : "Mark task complete",
        onToggle: async (nextDone) => {
          liveTask.done = nextDone;
          updateStatsDisplay();
          renderTaskList();
          await save();
        },
      });
      const content = el("div", { className: "work-item-content" });
      const textControl = createInlineWorkItemTextControl({
        textClassName: "task-text",
        value: liveTask.text || "",
        fallbackText: "Task",
        title: "Click to edit task",
        ariaLabel: "Edit task text",
        onCommit: async (nextText) => {
          liveTask.text = nextText;
          updateStatsDisplay();
          renderTaskList();
          await save();
        },
      });
      const attachments = createWorkItemAttachmentControls({
        kind: "task",
        owner: liveTask,
        deliverable,
        project,
        persistNow: true,
        scope: "projects-tab",
        onChange: () => {
          updateStatsDisplay();
        },
      });
      const pinBtn = createWorkItemPinButton({
        pinned: !!liveTask.pinned,
        titlePinned: "Unpin task",
        titleUnpinned: "Pin task",
        onToggle: async (nextPinned) => {
          liveTask.pinned = nextPinned;
          updateStatsDisplay();
          renderTaskList();
          await save();
        },
      });
      const deleteBtn = el("button", {
        className: "task-delete-btn",
        title: "Remove task",
        textContent: "×",
      });
      const actions = el("div", { className: "work-item-actions" });

      deleteBtn.addEventListener("click", async (e) => {
        e.stopPropagation();
        await clearAttachmentOwnerEmailRefs(liveTask);
        deliverable.tasks.splice(index, 1);
        updateStatsDisplay();
        renderTaskList();
        await save();
      });

      actions.append(attachments, pinBtn, deleteBtn);
      content.append(textControl);
      item.append(checkbox, content, actions);

      list.appendChild(item);
    });

    const addTaskRow = el("div", { className: "task-add-row" });
    const bulletSpan = el("span", {
      className: "task-icon task-add-bullet",
      textContent: "○",
    });
    const taskInput = el("input", {
      className: "task-add-input",
      type: "text",
      placeholder: "Add a task...",
    });

    let isCommittingTask = false;
    const commitPendingTask = async ({ refocus = false } = {}) => {
      const taskText = taskInput.value.trim();
      if (!taskText || isCommittingTask) return false;

      isCommittingTask = true;
      taskInput.value = "";
      try {
        deliverable.tasks.push({
          text: taskText,
          done: false,
          pinned: false,
          attachments: [],
        });

        updateStatsDisplay();
        await save();
        renderTaskList();

        if (refocus) {
          setTimeout(() => {
            const newTaskInput = list.querySelector(".task-add-input");
            if (newTaskInput) {
              newTaskInput.focus();
              newTaskInput.select();
            }
          }, 0);
        }
        return true;
      } finally {
        isCommittingTask = false;
      }
    };

    taskInput.addEventListener("keydown", async (e) => {
      if (e.key !== "Enter") return;
      e.preventDefault();
      e.stopPropagation();
      await commitPendingTask({ refocus: true });
    });

    taskInput.addEventListener("blur", () => {
      if (!taskInput.value.trim()) return;
      setTimeout(() => {
        void commitPendingTask({ refocus: false });
      }, 0);
    });

    taskInput.addEventListener("click", (e) => {
      e.stopPropagation();
    });

    addTaskRow.append(bulletSpan, taskInput);
    list.appendChild(addTaskRow);
  };

  renderTaskList();

  container.append(heading, list);
  return container;
}

function renderDeliverableCard(deliverable, isPrimary, project) {
  syncDeliverableWorkItemFields(deliverable);
  const deliverableId = String(deliverable?.id || createId("dlv")).trim();
  if (!deliverable?.id) deliverable.id = deliverableId;
  const card = el("div", {
    className: `deliverable-card-new ${isPrimary ? "is-primary" : ""} details-collapsed`,
  });
  card.dataset.deliverableId = deliverableId;

  const header = createCardHeader(deliverable, isPrimary, card, project);
  const progress = createProgressSection(deliverable);

  const statusSection = createDeliverableStatusSection(
    deliverable,
    project,
    card
  );

  const tasksPreview = createTasksPreview(deliverable, card, project);
  const notesSection = createNotesSection(deliverable, card, project);

  card.append(header, progress, statusSection, tasksPreview, notesSection);
  updateDeliverableWorkItemUi(card, deliverable);
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
  syncProjectActiveDeliverables(target, {
    fallbackActiveId: newDeliverable.id,
  });
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
        active: true,
      }));
    base.deliverables.sort(compareDeliverablesByDueDesc);
    if (!base.deliverables.length) base.deliverables = [createDeliverable()];
    syncProjectActiveDeliverables(base);
    autoSetPrimary(base);
    base.overviewSortDir = "desc";
    base.pinned = projects.some((project) => project?.pinned);
    base.pinnedOrder = base.pinned ? getLowestPinnedProjectOrder(projects) : null;

    db[sorted[0]] = base;
    for (let i = sorted.length - 1; i >= 1; i--) {
      db.splice(sorted[i], 1);
      removedCount++;
    }
    mergedCount++;
  });

  syncPinnedProjectOrders(db);
  save();
  render();
  toast(`Merged ${mergedCount} group${mergedCount === 1 ? "" : "s"} (${removedCount} removed).`);
}

function sortProjectsByCurrent(items, projectListContextMap = null) {
  items.sort((a, b) => {
    const dir = currentSort.dir === "asc" ? 1 : -1;
    if (currentSort.key === "due") {
      const bucketDiff = compareProjectListSortBuckets(a, b, projectListContextMap);
      if (bucketDiff) return bucketDiff;
      const da = getProjectSortKey(a, projectListContextMap?.get(a));
      const dbb = getProjectSortKey(b, projectListContextMap?.get(b));
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

function sortProjectsByDueDesc(items, projectListContextMap = null) {
  items.sort((a, b) => {
    const dir = currentSort.dir === "asc" ? 1 : -1;
    const bucketDiff = compareProjectListSortBuckets(a, b, projectListContextMap);
    if (bucketDiff) return bucketDiff;
    const da = getProjectSortKey(a, projectListContextMap?.get(a));
    const dbb = getProjectSortKey(b, projectListContextMap?.get(b));
    if (!da && !dbb) return 0;
    if (!da) return 1;
    if (!dbb) return -1;
    return (da - dbb) * dir;
  });
}

function compareDueDateValues(aDate, bDate, direction = "asc") {
  if (!aDate && !bDate) return 0;
  if (!aDate) return 1;
  if (!bDate) return -1;
  return direction === "desc" ? bDate - aDate : aDate - bDate;
}

function buildProjectDeliverableRowEntries(items, projectListContextMap = null) {
  const deliverableRows = [];
  items.forEach((project) => {
    const projectListContext = projectListContextMap?.get(project) || null;
    if (!projectListContext) return;
    const projectIndex = db.indexOf(project);
    projectListContext.visibleDeliverables.forEach((deliverable) => {
      deliverableRows.push({
        project,
        projectIndex,
        projectListContext,
        deliverable,
        dueDate: parseDueStr(deliverable?.due),
        isPinnedProject: !!project?.pinned,
        isCompleteDeliverable: isFinished(deliverable),
      });
    });
  });
  return deliverableRows;
}

function getProjectDeliverableRowSortBucket(row) {
  if (!shouldSortCompletedProjectsLast() || row?.isPinnedProject) return 0;
  if (!row?.isCompleteDeliverable) return row?.dueDate ? 0 : 1;
  return row?.dueDate ? 2 : 3;
}

function compareProjectDeliverableRowIdentity(a, b) {
  const projectIdDiff = String(a?.project?.id || "").localeCompare(
    String(b?.project?.id || ""),
    undefined,
    {
      numeric: true,
      sensitivity: "base",
    }
  );
  if (projectIdDiff) return projectIdDiff;

  const projectNameDiff = String(a?.project?.name || "").localeCompare(
    String(b?.project?.name || ""),
    undefined,
    {
      numeric: true,
      sensitivity: "base",
    }
  );
  if (projectNameDiff) return projectNameDiff;

  const deliverableNameDiff = String(a?.deliverable?.name || "").localeCompare(
    String(b?.deliverable?.name || ""),
    undefined,
    {
      numeric: true,
      sensitivity: "base",
    }
  );
  if (deliverableNameDiff) return deliverableNameDiff;

  return String(a?.deliverable?.id || "").localeCompare(
    String(b?.deliverable?.id || ""),
    undefined,
    {
      numeric: true,
      sensitivity: "base",
    }
  );
}

function compareProjectDeliverableRows(a, b) {
  const aPinned = !!a?.isPinnedProject;
  const bPinned = !!b?.isPinnedProject;
  if (aPinned !== bPinned) return aPinned ? -1 : 1;

  const dir = currentSort.dir === "asc" ? 1 : -1;
  if (currentSort.key !== "due") {
    const projectFieldDiff =
      String(a?.project?.[currentSort.key] || "").localeCompare(
        String(b?.project?.[currentSort.key] || ""),
        undefined,
        {
          numeric: true,
          sensitivity: "base",
        }
      ) * dir;
    if (projectFieldDiff) return projectFieldDiff;
    return compareProjectDeliverableRowIdentity(a, b);
  }

  if (!separateDeliverableCompletionGroups) {
    const mixedDueDiff = compareDueDateValues(
      a?.dueDate || null,
      b?.dueDate || null,
      "desc"
    );
    if (mixedDueDiff) return mixedDueDiff;
    return compareProjectDeliverableRowIdentity(a, b);
  }

  const sortBucketDiff =
    getProjectDeliverableRowSortBucket(a) - getProjectDeliverableRowSortBucket(b);
  if (sortBucketDiff) return sortBucketDiff;

  const dueDirection =
    shouldSortCompletedProjectsLast() &&
    !aPinned &&
    getProjectDeliverableRowSortBucket(a) >= 2
      ? "desc"
      : dir === -1
        ? "desc"
        : "asc";
  const dueDiff = compareDueDateValues(
    a?.dueDate || null,
    b?.dueDate || null,
    dueDirection
  );
  if (dueDiff) return dueDiff;

  return compareProjectDeliverableRowIdentity(a, b);
}

function sortProjectDeliverableRows(items) {
  items.sort((a, b) => compareProjectDeliverableRows(a, b));
}

function clearPinnedProjectDragStyles() {
  const tbody = document.getElementById("tbody");
  if (!tbody) return;
  tbody.querySelectorAll(".project-row").forEach((row) => {
    row.classList.remove("project-dragging", "project-drop-before", "project-drop-after");
    delete row.dataset.dropPosition;
  });
}

function enablePinnedProjectRowDrag(row, handle, project) {
  if (!row || !handle || !project?.pinned) return;
  handle.draggable = true;

  handle.addEventListener("dragstart", (e) => {
    pinnedProjectDragState = {
      project,
      didMove: false,
    };
    row.classList.add("project-dragging");
    if (e.dataTransfer) {
      e.dataTransfer.effectAllowed = "move";
      e.dataTransfer.setData(
        "text/plain",
        String(project?.id || project?.name || "pinned-project")
      );
    }
  });

  handle.addEventListener("dragend", () => {
    row.classList.remove("project-dragging");
    clearPinnedProjectDragStyles();
    if (pinnedProjectDragState?.project === project && pinnedProjectDragState.didMove) {
      projectPinHandleSuppressClickUntil = Date.now() + 250;
    }
    if (pinnedProjectDragState?.project === project) {
      pinnedProjectDragState = null;
    }
  });

  row.addEventListener("dragover", (e) => {
    if (!pinnedProjectDragState?.project?.pinned || pinnedProjectDragState.project === project) {
      return;
    }
    e.preventDefault();
    const rect = row.getBoundingClientRect();
    const before = e.clientY < rect.top + rect.height / 2;
    row.classList.toggle("project-drop-before", before);
    row.classList.toggle("project-drop-after", !before);
    row.dataset.dropPosition = before ? "before" : "after";
  });

  row.addEventListener("dragleave", () => {
    row.classList.remove("project-drop-before", "project-drop-after");
    delete row.dataset.dropPosition;
  });

  row.addEventListener("drop", async (e) => {
    if (!pinnedProjectDragState?.project?.pinned || pinnedProjectDragState.project === project) {
      return;
    }
    e.preventDefault();
    const before = row.dataset.dropPosition !== "after";
    const moved = movePinnedProjectToTarget(
      pinnedProjectDragState.project,
      project,
      before,
      db
    );
    clearPinnedProjectDragStyles();
    if (pinnedProjectDragState) {
      pinnedProjectDragState.didMove = moved;
    }
    pinnedProjectDragState = null;
    if (!moved) return;
    projectPinHandleSuppressClickUntil = Date.now() + 250;
    await save();
    renderProjectsPreservingExpandedDeliverables();
  });
}

function buildProjectTableRow(project, projectIndex, rowTemplate) {
  const tr = rowTemplate.content.cloneNode(true).querySelector("tr");
  tr.classList.add("project-row");

  const selectCell = tr.querySelector(".cell-select");
  const pinBtn = selectCell?.querySelector(".pin-btn");
  if (pinBtn) {
    const isPinned = !!project?.pinned;
    tr.classList.toggle("is-pinned-project", isPinned);
    pinBtn.classList.toggle("is-pinned", isPinned);
    pinBtn.setAttribute("aria-pressed", String(isPinned));
    pinBtn.setAttribute("aria-label", isPinned ? "Unpin project" : "Pin project");
    pinBtn.setAttribute("title", isPinned ? "Unpin project" : "Pin project");
    pinBtn.textContent = "";
    pinBtn.appendChild(createIcon(PIN_ICON_PATH, 14));
    pinBtn.draggable = false;
    if (isPinned) enablePinnedProjectRowDrag(tr, pinBtn, project);
    pinBtn.onclick = async (e) => {
      e.stopPropagation();
      if (Date.now() < projectPinHandleSuppressClickUntil) return;
      setProjectPinnedState(project, !project?.pinned, db);
      await save();
      renderProjectsPreservingExpandedDeliverables();
    };
  }

  const idCell = tr.querySelector(".cell-id");
  if (idCell) {
    const idBadge = idCell.querySelector(".id-badge") || idCell;
    idBadge.textContent = project?.id || "--";
  }

  const nameCell = tr.querySelector(".cell-name");
  if (nameCell) {
    nameCell.innerHTML = "";
    const projectDetailsHeader = el("div", {
      className: "project-details-header",
    });
    const projectDetailsMain = el("div", {
      className: "project-details-main",
    });

    if (project?.path) {
      const link = el("button", {
        className: "path-link",
        textContent: project?.name || "--",
        title: `Open: ${project.path}`,
      });
      link.onclick = async () => {
        try {
          await window.pywebview.api.open_path(convertPath(project.path));
          toast("Opening folder...");
        } catch (e) {
          toast("Failed to open path.");
        }
      };
      projectDetailsMain.appendChild(link);
    } else {
      projectDetailsMain.appendChild(
        el("span", {
          className: "project-title-text",
          textContent: project?.name || "--",
        })
      );
    }

    const account = extractAccountFromPath(project?.path);
    if (account) {
      projectDetailsMain.append(
        el("small", { className: "muted", textContent: ` (${account})` })
      );
    }

    if (project?.nick) {
      projectDetailsMain.append(
        el("small", { className: "muted", textContent: ` (${project.nick})` })
      );
    }

    const projectAttachmentInline = el("div", {
      className: "project-inline-attachment",
    });
    projectAttachmentInline.appendChild(
      createAttachmentControl(
        {
          kind: "project",
          owner: project,
          project,
          scope: "projects-tab",
        },
        {
          persistNow: true,
        }
      )
    );
    projectDetailsHeader.append(projectDetailsMain, projectAttachmentInline);
    nameCell.appendChild(projectDetailsHeader);

    const projectNotes = (project?.notes || "").trim();
    if (projectNotes) {
      nameCell.append(
        el("div", {
          className: "project-notes-snippet",
          textContent: projectNotes,
        })
      );
    }
  }

  const actionsCell = tr.querySelector(".cell-actions");
  if (actionsCell) {
    actionsCell.innerHTML = "";
    const actionsStack = el("div", { className: "actions-stack" });
    actionsStack.append(
      createIconButton({
        className: "btn icon-only",
        title: "Edit",
        ariaLabel: "Edit project",
        path: PENCIL_ICON_PATH,
        onClick: () => openEdit(projectIndex),
      }),
      createIconButton({
        className: "btn btn-danger icon-only",
        title: "Delete",
        ariaLabel: "Delete project",
        path: TRASH_ICON_PATH,
        onClick: () => removeProject(projectIndex),
      })
    );
    actionsCell.append(actionsStack);
  }

  return tr;
}

function renderGroupedProjectDeliverablesCell(
  deliverablesCell,
  project,
  projectListContext
) {
  if (!deliverablesCell) return;
  deliverablesCell.innerHTML = "";

  const visibleDeliverables = projectListContext.visibleDeliverables;
  const priorityId = projectListContext.activeAnchorDeliverable?.id;

  if (visibleDeliverables.length) {
    if (projectListContext.showTimeframeNote && projectListContext.timeframeNote) {
      deliverablesCell.appendChild(
        el("div", {
          className: "project-timeframe-note",
          textContent: projectListContext.timeframeNote,
        })
      );
    }

    const cardsContainer = el("div", { className: "deliverable-cards-container" });
    visibleDeliverables.forEach((deliverable) => {
      const isPrimary = deliverable.id === priorityId;
      cardsContainer.appendChild(renderDeliverableCard(deliverable, isPrimary, project));
    });

    deliverablesCell.appendChild(cardsContainer);
    return;
  }

  deliverablesCell.appendChild(
    el("div", {
      className: "deliverable-empty",
      textContent: "--",
    })
  );
}

function renderProjectDeliverableCell(
  deliverablesCell,
  deliverable,
  isPrimary,
  project
) {
  if (!deliverablesCell) return;
  deliverablesCell.innerHTML = "";
  if (!deliverable) {
    deliverablesCell.appendChild(
      el("div", {
        className: "deliverable-empty",
        textContent: "--",
      })
    );
    return;
  }
  deliverablesCell.appendChild(renderDeliverableCard(deliverable, isPrimary, project));
}

function appendProjectSearchContextRow(tbody, query, project, matchContextMap) {
  if (!query || !matchContextMap.has(project)) return;
  const contextRow = buildMatchContextRow(query, project, matchContextMap.get(project));
  if (contextRow) tbody.appendChild(contextRow);
}

function renderGroupedProjectRows({
  tbody,
  items,
  rowTemplate,
  projectListContextMap,
  matchContextMap,
  query,
  appendSectionSeparator,
}) {
  let lastWeekKey = null;
  let lastCompleteWeekKey = null;
  let pinnedSectionShown = false;
  let completeProjectsSectionShown = false;

  items.forEach((project) => {
    const projectListContext = projectListContextMap.get(project);
    if (!projectListContext) return;

    const isPinnedProject = !!project?.pinned;
    const projectDue = getProjectSortKey(project, projectListContext);
    const weekKey = projectDue ? formatWeekKey(projectDue) : "no-date";
    const isCompleteOnlyProject =
      shouldSortCompletedProjectsLast() && !projectListContext.hasIncompleteActiveWork;

    if (isPinnedProject) {
      if (!pinnedSectionShown) {
        appendSectionSeparator("Pinned Projects");
        pinnedSectionShown = true;
      }
    } else if (isCompleteOnlyProject) {
      if (!completeProjectsSectionShown) {
        appendSectionSeparator("Complete Projects");
        completeProjectsSectionShown = true;
      }
      if (lastCompleteWeekKey === null || weekKey !== lastCompleteWeekKey) {
        appendSectionSeparator(projectDue ? formatWeekDisplay(projectDue) : "");
      }
      lastCompleteWeekKey = weekKey;
    } else {
      if (lastWeekKey === null || weekKey !== lastWeekKey) {
        appendSectionSeparator(projectDue ? formatWeekDisplay(projectDue) : "");
      }
      lastWeekKey = weekKey;
    }

    const tr = buildProjectTableRow(project, db.indexOf(project), rowTemplate);
    renderGroupedProjectDeliverablesCell(
      tr.querySelector(".cell-deliverables"),
      project,
      projectListContext
    );
    tbody.appendChild(tr);
    appendProjectSearchContextRow(tbody, query, project, matchContextMap);
  });
}

function renderUngroupedDeliverableRows({
  tbody,
  items,
  rowTemplate,
  projectListContextMap,
  matchContextMap,
  query,
  appendSectionSeparator,
}) {
  const deliverableRows = buildProjectDeliverableRowEntries(
    items,
    projectListContextMap
  );
  sortProjectDeliverableRows(deliverableRows);
  const orderedDeliverableRows = [];
  if (items.some((project) => project?.pinned)) {
    const pinnedDeliverableRows = new Map();
    const unpinnedDeliverableRows = [];
    deliverableRows.forEach((row) => {
      if (row?.isPinnedProject) {
        const projectRows = pinnedDeliverableRows.get(row.project) || [];
        projectRows.push(row);
        pinnedDeliverableRows.set(row.project, projectRows);
      } else {
        unpinnedDeliverableRows.push(row);
      }
    });
    items.filter((project) => project?.pinned).forEach((project) => {
      const projectRows = pinnedDeliverableRows.get(project);
      if (projectRows?.length) orderedDeliverableRows.push(...projectRows);
    });
    orderedDeliverableRows.push(...unpinnedDeliverableRows);
  } else {
    orderedDeliverableRows.push(...deliverableRows);
  }

  let lastWeekKey = null;
  let lastCompleteWeekKey = null;
  let pinnedSectionShown = false;
  let completeDeliverablesSectionShown = false;
  const searchContextProjects = new Set();

  orderedDeliverableRows.forEach((row) => {
    const {
      project,
      projectIndex,
      projectListContext,
      deliverable,
      dueDate,
      isPinnedProject,
      isCompleteDeliverable,
    } = row;
    const weekKey = dueDate ? formatWeekKey(dueDate) : "no-date";
    const isSeparatedCompleteDeliverable =
      shouldSortCompletedProjectsLast() &&
      !isPinnedProject &&
      isCompleteDeliverable;

    if (isPinnedProject) {
      if (!pinnedSectionShown) {
        appendSectionSeparator("Pinned Projects");
        pinnedSectionShown = true;
      }
    } else if (isSeparatedCompleteDeliverable) {
      if (!completeDeliverablesSectionShown) {
        appendSectionSeparator("Complete Deliverables");
        completeDeliverablesSectionShown = true;
      }
      if (lastCompleteWeekKey === null || weekKey !== lastCompleteWeekKey) {
        appendSectionSeparator(dueDate ? formatWeekDisplay(dueDate) : "");
      }
      lastCompleteWeekKey = weekKey;
    } else {
      if (lastWeekKey === null || weekKey !== lastWeekKey) {
        appendSectionSeparator(dueDate ? formatWeekDisplay(dueDate) : "");
      }
      lastWeekKey = weekKey;
    }

    const tr = buildProjectTableRow(project, projectIndex, rowTemplate);
    const isPrimary =
      deliverable?.id === projectListContext.activeAnchorDeliverable?.id;
    renderProjectDeliverableCell(
      tr.querySelector(".cell-deliverables"),
      deliverable,
      isPrimary,
      project
    );
    tbody.appendChild(tr);

    if (!searchContextProjects.has(project)) {
      appendProjectSearchContextRow(tbody, query, project, matchContextMap);
      searchContextProjects.add(project);
    }
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

    dCtx.matchingNotes.forEach((noteText) => {
      const item = el("div", { className: "search-context-item" });
      item.append(
        el("span", { className: "search-context-label", textContent: d.name + " Note: " }),
        createContextSnippet(noteText, q)
      );
      container.appendChild(item);
    });

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
  syncProjectsFilterDropdowns();
  tbody.innerHTML = "";

  const q = val("search").toLowerCase();
  const matchContextMap = new Map();
  const projectListContextMap = new Map();

  let items = db.filter((p) => {
    if (q) {
      const ctx = getMatchContext(q, p);
      if (!ctx) return false;
      matchContextMap.set(p, ctx);
    }
    const projectListContext = getProjectListRenderContext(p);
    if (!projectListContext) return false;
    projectListContextMap.set(p, projectListContext);
    return projectListContext.matchesFilters;
  });

  const pinned = getPinnedProjectsInManualOrder(items);
  const unpinned = items.filter((p) => !p?.pinned);
  sortProjectsByCurrent(unpinned, projectListContextMap);
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

  const emptyTitle = emptyState?.querySelector("h3");
  const emptyBody = emptyState?.querySelector("p");
  const hasActiveProjectFilters =
    !!q ||
    dueFilter !== "all" ||
    statusFilter !== "all" ||
    deliverablesFilter === "incomplete";
  const emptyStateEntityLabel = groupDeliverablesByProject
    ? "projects"
    : "deliverables";
  if (emptyTitle) {
    emptyTitle.textContent = hasActiveProjectFilters
      ? `No matching ${emptyStateEntityLabel}`
      : `No ${emptyStateEntityLabel} yet`;
  }
  if (emptyBody) {
    emptyBody.textContent = hasActiveProjectFilters
      ? "Try a different search or adjust the current filters."
      : groupDeliverablesByProject
        ? "Create a new project to get started."
        : "Create a new project with a deliverable to get started.";
  }
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

  if (groupDeliverablesByProject) {
    renderGroupedProjectRows({
      tbody,
      items,
      rowTemplate,
      projectListContextMap,
      matchContextMap,
      query: q,
      appendSectionSeparator,
    });
    return;
  }

  renderUngroupedDeliverableRows({
    tbody,
    items,
    rowTemplate,
    projectListContextMap,
    matchContextMap,
    query: q,
    appendSectionSeparator,
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
    syncDeliverableNoteFields(d);
    const dCtx = {
      deliverable: d,
      fields: [],
      matchingTasks: [],
      matchingNotes: [],
    };
    if (str(d.name).includes(q)) dCtx.fields.push("name");
    if (str(d.due).includes(q)) dCtx.fields.push("due");
    (d.noteItems || []).forEach((noteItem) => {
      if (str(noteItem?.text).includes(q)) dCtx.matchingNotes.push(noteItem.text);
    });
    (d.tasks || []).forEach((t) => {
      if (str(t.text).includes(q)) dCtx.matchingTasks.push(t.text);
    });
    const hasStatusMatch = (d.statuses || []).some((s) => str(s).includes(q));
    if (
      dCtx.fields.length ||
      dCtx.matchingTasks.length ||
      dCtx.matchingNotes.length ||
      hasStatusMatch
    ) {
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
    localProjectPath: "",
    workroomRootPath: "",
    notes: "",
    refs: [],
    attachments: [],
    links: [],
    deliverables: [deliverable],
    overviewDeliverableId: deliverable.id,
    pinned: false,
    pinnedOrder: null,
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

  flushPendingModalDeliverableTasks();
  flushPendingModalDeliverableNotes();
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
  document.getElementById("f_path").title = p.path || "";

  const deliverableList = document.getElementById("deliverableList");
  deliverableList.innerHTML = "";
  setModalProjectDraft(p);
  const projectAttachmentHost = document.getElementById("modalProjectAttachmentHost");
  if (projectAttachmentHost) {
    projectAttachmentHost.replaceChildren(
      createAttachmentControl(
        {
          kind: "project",
          owner: getModalProjectDraft(),
          project: getModalProjectDraft(),
          scope: "edit-modal",
        },
        {
          persistNow: false,
        }
      )
    );
  }
  const sortedDeliverables = p.deliverables.slice();
  const activeAnchorDeliverable = getActiveAnchorDeliverable(p);
  sortDeliverablesByPrimaryThenDueDesc(
    sortedDeliverables,
    activeAnchorDeliverable?.id
  );
  sortedDeliverables.forEach((deliverable) =>
    addDeliverableCard(deliverable, activeAnchorDeliverable?.id, {
      projectDraft: p,
    })
  );
  if (!deliverableList.children.length) {
    addDeliverableCard(createDeliverable(), activeAnchorDeliverable?.id, {
      projectDraft: p,
    });
  }
  ensureModalProjectHasActiveDeliverable();

  document.getElementById("refList").innerHTML = "";
  (p.refs || []).forEach(addRefRowFrom);
}

function getDeliverableCardEmailRefs(card) {
  return buildLegacyEmailRefsFromAttachments(getDeliverableCardAttachments(card));
}

function setDeliverableCardEmailRefs(card, emailRefs) {
  if (!card) return;
  const attachments = normalizeAttachments(getDeliverableCardAttachments(card), {
    legacyEmailRefs: emailRefs,
  }).filter((attachment) => attachment.type !== "email");
  normalizeEmailRefs(emailRefs).forEach((emailRef) => {
    attachments.push(
      normalizeAttachmentEntry({
        type: "email",
        emailRef,
      })
    );
  });
  setDeliverableCardAttachments(card, attachments);
}

function getDeliverableCardLinks(card) {
  return buildLegacyLinksFromAttachments(getDeliverableCardAttachments(card));
}

function setDeliverableCardLinks(card, links) {
  if (!card) return;
  const attachments = normalizeAttachments(getDeliverableCardAttachments(card), {
    legacyLinks: links,
  }).filter((attachment) => attachment.type === "email");
  normalizeDeliverableLinks(links).forEach((link) => {
    attachments.push(
      normalizeAttachmentEntry({
        type: isWebAttachmentTarget(link.raw || link.url || "") ? "url" : "path",
        description: link.label || "",
        target: link.raw || link.url || "",
      })
    );
  });
  setDeliverableCardAttachments(card, attachments);
}

function getDeliverableCardAttachments(card) {
  if (!card) return [];
  if (!Array.isArray(card._attachmentItems)) {
    let parsed = null;
    try {
      parsed = JSON.parse(card.dataset.attachments || "null");
    } catch {
      parsed = null;
    }
    card._attachmentItems = normalizeAttachments(parsed, {
      legacyLinks: (() => {
        try {
          return JSON.parse(card.dataset.links || "[]");
        } catch {
          return [];
        }
      })(),
      legacyEmailRefs: (() => {
        try {
          return JSON.parse(card.dataset.emailRefs || "[]");
        } catch {
          return [];
        }
      })(),
      legacyEmailRef: (() => {
        const raw = card.dataset.emailRef || "";
        if (!raw) return null;
        try {
          return JSON.parse(raw);
        } catch {
          return raw;
        }
      })(),
    });
  }
  return normalizeAttachments(card._attachmentItems);
}

function setDeliverableCardAttachments(card, attachments) {
  if (!card) return [];
  const normalized = normalizeAttachments(attachments);
  card._attachmentItems = normalized;
  const compatOwner = { attachments: normalized };
  syncAttachmentOwnerCompatFields(compatOwner, {
    includeEmailRefs: true,
  });
  if (normalized.length) {
    card.dataset.attachments = JSON.stringify(normalized);
  } else {
    delete card.dataset.attachments;
  }
  if (compatOwner.links.length) {
    card.dataset.links = JSON.stringify(compatOwner.links);
  } else {
    delete card.dataset.links;
  }
  if (compatOwner.emailRefs.length) {
    card.dataset.emailRefs = JSON.stringify(compatOwner.emailRefs);
    card.dataset.emailRef = JSON.stringify(compatOwner.emailRefs[0]);
  } else {
    delete card.dataset.emailRefs;
    delete card.dataset.emailRef;
  }
  return normalized;
}

function getModalProjectDraft() {
  const list = document.getElementById("deliverableList");
  if (!list) return null;
  if (!list._projectDraft || typeof list._projectDraft !== "object") {
    list._projectDraft = { attachments: [], links: [] };
  }
  syncProjectAttachmentFields(list._projectDraft);
  return list._projectDraft;
}

function setModalProjectDraft(project) {
  const list = document.getElementById("deliverableList");
  if (!list) return null;
  list._projectDraft =
    project && typeof project === "object"
      ? project
      : { attachments: [], links: [] };
  syncProjectAttachmentFields(list._projectDraft);
  return list._projectDraft;
}

function getModalProjectLinks() {
  return buildLegacyLinksFromAttachments(getModalProjectAttachments());
}

function setModalProjectLinks(links) {
  const draft = getModalProjectDraft();
  if (!draft) return [];
  const attachments = normalizeAttachments(draft.attachments, {
    legacyLinks: links,
  }).filter((attachment) => attachment.type === "email");
  normalizeDeliverableLinks(links).forEach((link) => {
    attachments.push(
      normalizeAttachmentEntry({
        type: isWebAttachmentTarget(link.raw || link.url || "") ? "url" : "path",
        description: link.label || "",
        target: link.raw || link.url || "",
      })
    );
  });
  draft.attachments = normalizeAttachments(attachments);
  syncProjectAttachmentFields(draft);
  return draft.links;
}

function getModalProjectAttachments() {
  const draft = getModalProjectDraft();
  if (!draft) return [];
  syncProjectAttachmentFields(draft);
  return normalizeAttachments(draft.attachments);
}

function setModalProjectAttachments(attachments) {
  const draft = getModalProjectDraft();
  if (!draft) return [];
  draft.attachments = normalizeAttachments(attachments);
  syncProjectAttachmentFields(draft);
  return draft.attachments;
}

function getModalActiveDeliverableInputs(list = document.getElementById("deliverableList")) {
  if (!list) return [];
  return Array.from(list.querySelectorAll(".d-active"));
}

function getModalFallbackActiveDeliverableInput(list, excludedCard = null) {
  const activeInputs = getModalActiveDeliverableInputs(list);
  if (!activeInputs.length) return null;
  if (!excludedCard) return activeInputs[0];
  const remainingInputs = activeInputs.filter(
    (input) => input.closest(".deliverable-card") !== excludedCard
  );
  return remainingInputs[0] || activeInputs[0];
}

function ensureModalProjectHasActiveDeliverable({ preferredCard = null } = {}) {
  const list = document.getElementById("deliverableList");
  if (!list) return null;
  const checkedInputs = Array.from(list.querySelectorAll(".d-active:checked"));
  if (checkedInputs.length) return checkedInputs[0];

  const preferredInput =
    preferredCard && list.contains(preferredCard)
      ? preferredCard.querySelector(".d-active")
      : null;
  const fallbackInput =
    preferredInput || getModalFallbackActiveDeliverableInput(list);
  if (fallbackInput) fallbackInput.checked = true;
  return fallbackInput;
}

function addDeliverableCard(deliverable, activeAnchorId, options = {}) {
  const list = document.getElementById("deliverableList");
  const template = document.getElementById("deliverable-card-template");
  if (!list || !template) return;
  const { insertAtTop = false, projectDraft = getModalProjectDraft() } = options;

  const card = template.content
    .cloneNode(true)
    .querySelector(".deliverable-card");
  const deliverableId = deliverable.id || createId("dlv");
  card.dataset.deliverableId = deliverableId;
  syncDeliverableAttachmentFields(deliverable);
  setDeliverableCardAttachments(card, deliverable.attachments);

  card.querySelector(".d-name").value = deliverable.name || "";
  card.querySelector(".d-due").value = deliverable.due || "";

  const activeInput = card.querySelector(".d-active");
  if (deliverable.active === true || (activeAnchorId && deliverableId === activeAnchorId)) {
    activeInput.checked = true;
  }
  activeInput.addEventListener("change", () => {
    if (activeInput.checked) return;
    const fallbackInput = getModalFallbackActiveDeliverableInput(list, card);
    if (!list.querySelector(".d-active:checked") && fallbackInput) {
      fallbackInput.checked = true;
    }
  });

  const attachmentControlHost = card.querySelector(".deliverable-attachment-control");
  if (attachmentControlHost) {
    attachmentControlHost.replaceChildren(
      createAttachmentControl(
        {
          kind: "deliverable",
          owner: deliverable,
          deliverable,
          project: projectDraft,
          modalCard: card,
          scope: "edit-modal",
        },
        {
          persistNow: false,
        }
      )
    );
  }

  const taskList = card.querySelector(".deliverable-task-list");
  const noteList = card.querySelector(".deliverable-note-list");
  card._taskItems = (deliverable.tasks || []).map(normalizeTask);
  card._noteItems = normalizeDeliverableNoteItems(
    deliverable.noteItems,
    deliverable.notes || ""
  );
  renderModalDeliverableTaskList(card, taskList);
  renderModalDeliverableNoteList(card, noteList);
  card.querySelector(".btn-remove").onclick = async () => {
    const cardRefs = buildLegacyEmailRefsFromAttachments(getDeliverableCardAttachments(card));
    const workItemRefs = getModalDeliverableWorkItemEmailRefs(card);
    await stageModalManagedEmailRefsForRemoval(
      cardRefs.length ? cardRefs : deliverable.emailRefs
    );
    await stageModalManagedEmailRefsForRemoval(workItemRefs);
    if (openAttachmentPanelContext?.modalCard === card) {
      void requestAttachmentPanelClose();
    }
    card.remove();
    if (!list.querySelector(".deliverable-card")) {
      addDeliverableCard(createDeliverable(), null, { projectDraft });
    } else {
      ensureModalProjectHasActiveDeliverable();
    }
  };

  const statusContainer = card.querySelector(".deliverable-status");
  const markTasksDone = () => {
    getModalDeliverableTaskItems(card).forEach((task) => {
      task.done = true;
    });
    renderModalDeliverableTaskList(card, taskList);
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

  ensureModalProjectHasActiveDeliverable({ preferredCard: card });

  return card;
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

function getModalDeliverableTaskItems(card) {
  if (!card) return [];
  if (!Array.isArray(card._taskItems)) {
    card._taskItems = [];
  }
  return card._taskItems;
}

function getModalDeliverableNoteItems(card) {
  if (!card) return [];
  if (!Array.isArray(card._noteItems)) {
    card._noteItems = [];
  }
  return card._noteItems;
}

function collectWorkItemEmailRefs(items = [], normalizeItem = (item) => item) {
  const refs = [];
  (Array.isArray(items) ? items : []).forEach((item) => {
    const normalizedItem = normalizeItem(item);
    normalizeEmailRefs(normalizedItem?.emailRefs).forEach((emailRef) => {
      refs.push(emailRef);
    });
  });
  return refs;
}

function getModalDeliverableWorkItemEmailRefs(card) {
  return [
    ...collectWorkItemEmailRefs(getModalDeliverableTaskItems(card), normalizeTask),
    ...collectWorkItemEmailRefs(
      getModalDeliverableNoteItems(card),
      normalizeDeliverableNoteItem
    ),
  ];
}

function scrollElementIntoNearestModalBody(element) {
  if (!element) return;
  requestAnimationFrame(() => {
    const modalBody = element.closest(".modal-body");
    if (!modalBody) return;
    const elementRect = element.getBoundingClientRect();
    const modalBodyRect = modalBody.getBoundingClientRect();
    if (
      elementRect.top < modalBodyRect.top ||
      elementRect.bottom > modalBodyRect.bottom
    ) {
      element.scrollIntoView({
        behavior: "smooth",
        block: "nearest",
        inline: "nearest",
      });
    }
  });
}

function focusModalTaskInput(container) {
  const input = container?.querySelector(".task-add-input");
  if (!input) return;
  input.focus();
  input.select();
  scrollElementIntoNearestModalBody(input);
}

function focusModalNoteInput(container) {
  const input = container?.querySelector(".note-add-input");
  if (!input) return;
  input.focus();
  input.select();
  scrollElementIntoNearestModalBody(input);
}

function createWorkItemPinButton({
  pinned = false,
  titlePinned = "Unpin item",
  titleUnpinned = "Pin item",
  className = "",
  onToggle = null,
} = {}) {
  let currentPinned = !!pinned;
  let isTogglePending = false;
  const button = el("button", {
    className: `work-item-pin-btn ${className}`.trim(),
    type: "button",
  });
  button.appendChild(createIcon(PIN_ICON_PATH, 12));

  const syncState = () => {
    button.classList.toggle("is-pinned", currentPinned);
    button.title = currentPinned ? titlePinned : titleUnpinned;
    button.setAttribute(
      "aria-label",
      currentPinned ? titlePinned : titleUnpinned
    );
    button.setAttribute("aria-pressed", String(currentPinned));
    button.setAttribute("aria-busy", String(isTogglePending));
  };

  button.addEventListener("click", async (e) => {
    e.stopPropagation();
    if (isTogglePending) return;
    const previousPinned = currentPinned;
    const nextPinned = !currentPinned;
    currentPinned = nextPinned;
    isTogglePending = true;
    syncState();
    try {
      const resolvedPinned = onToggle ? await onToggle(nextPinned) : nextPinned;
      if (typeof resolvedPinned === "boolean") {
        currentPinned = resolvedPinned;
      }
    } catch (error) {
      currentPinned = previousPinned;
      console.warn("Failed to toggle work item pin state:", error);
    } finally {
      isTogglePending = false;
      syncState();
    }
  });

  syncState();
  return button;
}

function getModalDeliverableCardContext(card) {
  if (!card) return null;
  return {
    id: String(card.dataset.deliverableId || "").trim(),
    name: String(card.querySelector(".d-name")?.value || "").trim(),
  };
}

function createWorkItemAttachmentControls({
  kind = "task",
  owner = null,
  deliverable = null,
  project = null,
  modalCard = null,
  persistNow = false,
  scope = "projects-tab",
  onChange = null,
} = {}) {
  const container = el("div", { className: "work-item-attachment-control" });
  if (!owner) return container;
  container.appendChild(
    createAttachmentControl(
      {
        kind,
        owner,
        deliverable,
        project,
        modalCard,
        scope,
      },
      {
        persistNow,
        onChange,
        className: "is-work-item",
      }
    )
  );
  return container;
}

function createPinnedAttachmentPreviewButton(attachment) {
  const normalized = normalizeAttachmentEntry(attachment);
  if (!normalized) return null;
  const label =
    String(
      normalized.description ||
        getAttachmentDescriptionFallback(normalized) ||
        "Attachment"
    ).trim() || "Attachment";
  const button = el("button", {
    className: "deliverable-pinned-link",
    type: "button",
    textContent: label,
    title:
      normalized.type === "email"
        ? label
        : normalized.target || label,
  });
  button.addEventListener("click", async (e) => {
    e.stopPropagation();
    await openAttachmentEntry(normalized);
  });
  return button;
}

function createPinnedWorkItemAttachments(previewItem = {}) {
  const attachments = normalizeAttachments(previewItem.attachments, {
    legacyLinks: previewItem.links,
    legacyEmailRefs: previewItem.emailRefs,
  });
  if (!attachments.length) return null;

  const container = el("div", {
    className: "deliverable-pinned-item-attachments",
  });
  attachments.forEach((attachment) => {
    const button = createPinnedAttachmentPreviewButton(attachment);
    if (button) container.appendChild(button);
  });
  return container.childElementCount ? container : null;
}

function renderModalDeliverableTaskList(card, container, options = {}) {
  if (!card || !container) return;
  const { focusInput = false, ensureInputVisible = false } = options;
  const taskItems = getModalDeliverableTaskItems(card);
  container.innerHTML = "";

  taskItems.forEach((task, index) => {
    const item = el("div", {
      className: `deliverable-task-item ${task.done ? "done" : "undone"}`,
    });
    const iconSpan = el("span", {
      className: "task-icon",
      textContent: task.done ? "✓" : "○",
    });
    const textSpan = el("span", {
      className: "task-text",
      textContent: task.text || "Task",
    });
    const deleteBtn = el("button", {
      className: "task-delete-btn",
      type: "button",
      title: "Remove task",
      textContent: "×",
    });

    deleteBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      taskItems.splice(index, 1);
      renderModalDeliverableTaskList(card, container);
    });

    item.append(iconSpan, textSpan, deleteBtn);
    item.addEventListener("click", (e) => {
      if (e.target === deleteBtn) return;
      e.stopPropagation();
      task.done = !task.done;
      renderModalDeliverableTaskList(card, container);
    });
    container.appendChild(item);
  });

  const addTaskRow = el("div", { className: "task-add-row modal-task-add-row" });
  const bulletSpan = el("span", {
    className: "task-icon task-add-bullet",
    textContent: "○",
  });
  const taskInput = el("input", {
    className: "task-add-input",
    type: "text",
    placeholder: "Add a task...",
  });

  taskInput.addEventListener("keydown", (e) => {
    if (e.key !== "Enter") return;
    e.preventDefault();
    e.stopPropagation();
    commitPendingModalDeliverableTask(card, container, {
      refocus: true,
      rerender: true,
      ensureInputVisible: true,
    });
  });
  taskInput.addEventListener("blur", () => {
    if (!taskInput.value.trim()) return;
    setTimeout(() => {
      commitPendingModalDeliverableTask(card, container, {
        refocus: false,
        rerender: true,
        ensureInputVisible: true,
      });
    }, 0);
  });
  taskInput.addEventListener("click", (e) => {
    e.stopPropagation();
  });

  addTaskRow.append(bulletSpan, taskInput);
  container.appendChild(addTaskRow);

  if (focusInput) {
    setTimeout(() => {
      focusModalTaskInput(container);
    }, 0);
  } else if (ensureInputVisible) {
    scrollElementIntoNearestModalBody(taskInput);
  }
}

function commitPendingModalDeliverableTask(card, container, options = {}) {
  const {
    refocus = false,
    rerender = true,
    ensureInputVisible = true,
  } = options;
  const taskInput = container?.querySelector(".task-add-input");
  if (!card || !container || !taskInput) return false;

  const taskText = taskInput.value.trim();
  if (!taskText) return false;

  taskInput.value = "";
  const taskItems = getModalDeliverableTaskItems(card);
  taskItems.push({ text: taskText, done: false, links: [] });

  if (rerender) {
    renderModalDeliverableTaskList(card, container, {
      focusInput: refocus,
      ensureInputVisible,
    });
  }
  return true;
}

function flushPendingModalDeliverableTasks() {
  document.querySelectorAll("#deliverableList .deliverable-card").forEach((card) => {
    const taskList = card.querySelector(".deliverable-task-list");
    commitPendingModalDeliverableTask(card, taskList, {
      refocus: false,
      rerender: false,
      ensureInputVisible: false,
    });
  });
}

function renderModalDeliverableTaskList(card, container, options = {}) {
  if (!card || !container) return;
  const { focusInput = false, ensureInputVisible = false } = options;
  const taskItems = getModalDeliverableTaskItems(card);
  const modalDeliverable = getModalDeliverableCardContext(card);
  container.innerHTML = "";

  getPinnedFirstEntries(taskItems, normalizeTask).forEach(({ item: task, index }) => {
    taskItems[index] = normalizeTask(taskItems[index]);
    const liveTask = taskItems[index];
    const item = el("div", {
      className: `deliverable-task-item ${liveTask.done ? "done" : "undone"}`,
    });
    const checkbox = createWorkItemDoneCheckbox({
      checked: !!liveTask.done,
      ariaLabel: liveTask.done ? "Mark task incomplete" : "Mark task complete",
      onToggle: async (nextDone) => {
        liveTask.done = nextDone;
        renderModalDeliverableTaskList(card, container);
      },
    });
    const content = el("div", { className: "work-item-content" });
    const textControl = createInlineWorkItemTextControl({
      textClassName: "task-text",
      value: liveTask.text || "",
      fallbackText: "Task",
      title: "Click to edit task",
      ariaLabel: "Edit task text",
      onCommit: async (nextText) => {
        liveTask.text = nextText;
        renderModalDeliverableTaskList(card, container);
      },
    });
    const attachments = createWorkItemAttachmentControls({
      kind: "task",
      owner: liveTask,
      deliverable: modalDeliverable,
      project: getModalProjectDraft(),
      modalCard: card,
      persistNow: false,
      scope: "edit-modal",
    });
    const pinBtn = createWorkItemPinButton({
      pinned: !!liveTask.pinned,
      titlePinned: "Unpin task",
      titleUnpinned: "Pin task",
      onToggle: (nextPinned) => {
        liveTask.pinned = nextPinned;
        renderModalDeliverableTaskList(card, container);
      },
    });
    const deleteBtn = el("button", {
      className: "task-delete-btn",
      type: "button",
      title: "Remove task",
      textContent: "×",
    });
    const actions = el("div", { className: "work-item-actions" });

    deleteBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      void (async () => {
        await clearAttachmentOwnerEmailRefs(liveTask, { modalCard: card });
        taskItems.splice(index, 1);
        renderModalDeliverableTaskList(card, container);
      })();
    });

    actions.append(attachments, pinBtn, deleteBtn);
    content.append(textControl);
    item.append(checkbox, content, actions);
    container.appendChild(item);
  });

  const addTaskRow = el("div", { className: "task-add-row modal-task-add-row" });
  const bulletSpan = el("span", {
    className: "task-icon task-add-bullet",
    textContent: "○",
  });
  const taskInput = el("input", {
    className: "task-add-input",
    type: "text",
    placeholder: "Add a task...",
  });

  taskInput.addEventListener("keydown", (e) => {
    if (e.key !== "Enter") return;
    e.preventDefault();
    e.stopPropagation();
    commitPendingModalDeliverableTask(card, container, {
      refocus: true,
      rerender: true,
      ensureInputVisible: true,
    });
  });
  taskInput.addEventListener("blur", () => {
    if (!taskInput.value.trim()) return;
    setTimeout(() => {
      commitPendingModalDeliverableTask(card, container, {
        refocus: false,
        rerender: true,
        ensureInputVisible: true,
      });
    }, 0);
  });
  taskInput.addEventListener("click", (e) => {
    e.stopPropagation();
  });

  addTaskRow.append(bulletSpan, taskInput);
  container.appendChild(addTaskRow);

  if (focusInput) {
    setTimeout(() => {
      focusModalTaskInput(container);
    }, 0);
  } else if (ensureInputVisible) {
    scrollElementIntoNearestModalBody(taskInput);
  }
}

function commitPendingModalDeliverableTask(card, container, options = {}) {
  const {
    refocus = false,
    rerender = true,
    ensureInputVisible = true,
  } = options;
  const taskInput = container?.querySelector(".task-add-input");
  if (!card || !container || !taskInput) return false;

  const taskText = taskInput.value.trim();
  if (!taskText) return false;

  taskInput.value = "";
  const taskItems = getModalDeliverableTaskItems(card);
  taskItems.push({
    text: taskText,
    done: false,
    pinned: false,
    attachments: [],
  });

  if (rerender) {
    renderModalDeliverableTaskList(card, container, {
      focusInput: refocus,
      ensureInputVisible,
    });
  }
  return true;
}

function flushPendingModalDeliverableTasks() {
  document.querySelectorAll("#deliverableList .deliverable-card").forEach((card) => {
    const taskList = card.querySelector(".deliverable-task-list");
    commitPendingModalDeliverableTask(card, taskList, {
      refocus: false,
      rerender: false,
      ensureInputVisible: false,
    });
  });
}

function renderModalDeliverableNoteList(card, container, options = {}) {
  if (!card || !container) return;
  const { focusInput = false, ensureInputVisible = false } = options;
  const noteItems = getModalDeliverableNoteItems(card);
  const modalDeliverable = getModalDeliverableCardContext(card);
  container.innerHTML = "";

  getPinnedFirstEntries(noteItems, normalizeDeliverableNoteItem).forEach(
    ({ item: noteItem, index }) => {
      noteItems[index] = normalizeDeliverableNoteItem(noteItems[index]);
      const liveNote = noteItems[index];
      const item = el("div", { className: "deliverable-note-item" });
      const iconSpan = el("span", {
        className: "note-icon",
        "aria-hidden": "true",
      });
      iconSpan.appendChild(createIcon(NOTE_ICON_PATH, 12));
      const content = el("div", { className: "work-item-content" });
      const textControl = createInlineWorkItemTextControl({
        textClassName: "note-text",
        value: liveNote.text || "",
        fallbackText: "Note",
        title: "Click to edit note",
        ariaLabel: "Edit note text",
        onCommit: async (nextText) => {
          liveNote.text = nextText;
          renderModalDeliverableNoteList(card, container);
        },
      });
      const attachments = createWorkItemAttachmentControls({
        kind: "note",
        owner: liveNote,
        deliverable: modalDeliverable,
        project: getModalProjectDraft(),
        modalCard: card,
        persistNow: false,
        scope: "edit-modal",
      });
      const pinBtn = createWorkItemPinButton({
        pinned: !!liveNote.pinned,
        titlePinned: "Unpin note",
        titleUnpinned: "Pin note",
        onToggle: (nextPinned) => {
          liveNote.pinned = nextPinned;
          renderModalDeliverableNoteList(card, container);
          return nextPinned;
        },
      });
      const deleteBtn = el("button", {
        className: "task-delete-btn",
        type: "button",
        title: "Remove note",
        textContent: "×",
      });
      const actions = el("div", { className: "work-item-actions" });

      deleteBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        void (async () => {
          await clearAttachmentOwnerEmailRefs(liveNote, { modalCard: card });
          noteItems.splice(index, 1);
          renderModalDeliverableNoteList(card, container);
        })();
      });

      actions.append(attachments, pinBtn, deleteBtn);
      content.append(textControl);
      item.append(iconSpan, content, actions);
      container.appendChild(item);
    }
  );

  const addNoteRow = el("div", { className: "task-add-row modal-task-add-row" });
  const noteBullet = el("span", {
    className: "note-icon note-add-bullet",
    "aria-hidden": "true",
  });
  noteBullet.appendChild(createIcon(NOTE_ICON_PATH, 12));
  const noteInput = el("input", {
    className: "task-add-input note-add-input",
    type: "text",
    placeholder: "Add a note...",
  });

  noteInput.addEventListener("keydown", (e) => {
    if (e.key !== "Enter") return;
    e.preventDefault();
    e.stopPropagation();
    commitPendingModalDeliverableNote(card, container, {
      refocus: true,
      rerender: true,
      ensureInputVisible: true,
    });
  });
  noteInput.addEventListener("blur", () => {
    if (!noteInput.value.trim()) return;
    setTimeout(() => {
      commitPendingModalDeliverableNote(card, container, {
        refocus: false,
        rerender: true,
        ensureInputVisible: true,
      });
    }, 0);
  });
  noteInput.addEventListener("click", (e) => {
    e.stopPropagation();
  });

  addNoteRow.append(noteBullet, noteInput);
  container.appendChild(addNoteRow);

  if (focusInput) {
    setTimeout(() => {
      focusModalNoteInput(container);
    }, 0);
  } else if (ensureInputVisible) {
    scrollElementIntoNearestModalBody(noteInput);
  }
}

function commitPendingModalDeliverableNote(card, container, options = {}) {
  const {
    refocus = false,
    rerender = true,
    ensureInputVisible = true,
  } = options;
  const noteInput = container?.querySelector(".note-add-input");
  if (!card || !container || !noteInput) return false;

  const noteText = noteInput.value.trim();
  if (!noteText) return false;

  noteInput.value = "";
  const noteItems = getModalDeliverableNoteItems(card);
  noteItems.push({
    text: noteText,
    pinned: false,
    attachments: [],
  });

  if (rerender) {
    renderModalDeliverableNoteList(card, container, {
      focusInput: refocus,
      ensureInputVisible,
    });
  }
  return true;
}

function flushPendingModalDeliverableNotes() {
  document.querySelectorAll("#deliverableList .deliverable-card").forEach((card) => {
    const noteList = card.querySelector(".deliverable-note-list");
    commitPendingModalDeliverableNote(card, noteList, {
      refocus: false,
      rerender: false,
      ensureInputVisible: false,
    });
  });
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
    path: normalizeProjectPath(val("f_path")),
    localProjectPath: normalizeWindowsPath(existingProject?.localProjectPath || ""),
    workroomRootPath: normalizeWindowsPath(existingProject?.workroomRootPath || ""),
    refs: [],
    attachments: getModalProjectAttachments(),
    links: getModalProjectLinks(),
    deliverables: [],
    overviewDeliverableId: "",
    pinned: !!existingProject?.pinned,
    pinnedOrder: normalizePinnedProjectOrder(existingProject?.pinnedOrder),
    lightingSchedule,
    title24,
  };

  document.querySelectorAll("#deliverableList .deliverable-card").forEach((card) => {
    const deliverableId = card.dataset.deliverableId || createId("dlv");
    const name = card.querySelector(".d-name").value.trim();
    const due = card.querySelector(".d-due").value.trim();
    const statuses = readStatusPickerFrom(
      card.querySelector(".deliverable-status")
    );
    const attachments = getDeliverableCardAttachments(card);

    const tasks = getModalDeliverableTaskItems(card)
      .map((task) => {
        const normalizedTask = normalizeTask(task);
        const text = String(normalizedTask.text || "").trim();
        if (!text) return null;
        return {
          text,
          done: !!normalizedTask.done,
          pinned: !!normalizedTask.pinned,
          attachments: normalizedTask.attachments,
          links: normalizedTask.links,
          emailRefs: normalizedTask.emailRefs,
        };
      })
      .filter(Boolean);
    const noteItems = getModalDeliverableNoteItems(card)
      .map((noteItem) => {
        const normalizedNoteItem = normalizeDeliverableNoteItem(noteItem);
        const text = String(normalizedNoteItem.text || "").trim();
        if (!text) return null;
        return {
          text,
          pinned: !!normalizedNoteItem.pinned,
          attachments: normalizedNoteItem.attachments,
          links: normalizedNoteItem.links,
          emailRefs: normalizedNoteItem.emailRefs,
        };
      })
      .filter(Boolean);

    const deliverable = normalizeDeliverable({
      id: deliverableId,
      name,
      due,
      noteItems,
      tasks,
      statuses,
      attachments,
      active: !!card.querySelector(".d-active")?.checked,
    });

    out.deliverables.push(deliverable);
  });

  if (!out.deliverables.length) {
    out.deliverables = [createDeliverable()];
  }

  document.querySelectorAll("#refList .ref-row").forEach((row) => {
    const label = row.querySelector(".r-label").value.trim();
    const raw = row.querySelector(".r-url").value.trim();
    if (!raw) return;
    const link = normalizeLink(raw);
    if (label) link.label = label;
    out.refs.push(link);
  });
  return normalizeProject(out);
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
  addDeliverableCard(createDeliverable(), null, {
    insertAtTop: true,
    projectDraft: getModalProjectDraft(),
  });
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

function normalizeNotebookType(type) {
  return type === "checklist" ? "checklist" : "note";
}

function hasActiveNoteSelection() {
  return Boolean(activeNoteTab && noteTabs.includes(activeNoteTab));
}

function hasActiveChecklistSelection() {
  return Boolean(getChecklistById(activeChecklistTabId));
}

function ensureActiveNotebookSelection(preferredType = activeNotebookType) {
  const normalizedPreferred = normalizeNotebookType(preferredType);
  const hasNote = hasActiveNoteSelection();
  const hasChecklist = hasActiveChecklistSelection();

  if (normalizedPreferred === "note" && hasNote) {
    activeNotebookType = "note";
    return;
  }

  if (normalizedPreferred === "checklist" && hasChecklist) {
    activeNotebookType = "checklist";
    return;
  }

  if (hasNote) {
    activeNotebookType = "note";
    return;
  }

  if (hasChecklist) {
    activeNotebookType = "checklist";
    return;
  }

  activeNotebookType = "note";
}

function syncNotebookSearchUi() {
  const searchInput = document.getElementById("notesSearch");
  if (!searchInput) return;

  const isChecklistMode = activeNotebookType === "checklist";
  const hasActiveSelection = isChecklistMode
    ? hasActiveChecklistSelection()
    : hasActiveNoteSelection();
  const nextValue = isChecklistMode ? checklistSearchQuery : notesSearchQuery;

  searchInput.placeholder = isChecklistMode
    ? "Search checklist items..."
    : "Search notes...";
  searchInput.disabled = !hasActiveSelection;
  if (searchInput.value !== nextValue) {
    searchInput.value = nextValue;
  }
}

function handleNotebookSearchInput(e) {
  const nextValue = String(e?.target?.value || "");
  if (activeNotebookType === "checklist") {
    checklistSearchQuery = nextValue;
    renderChecklistSearchResults();
    return;
  }

  notesSearchQuery = nextValue;
  renderNoteSearchResults();
}

function renderTextsView() {
  ensureActiveNotebookSelection();

  const searchWrap = document.getElementById("textsSearchWrap");
  const notesPane = document.getElementById("textsNotesPane");
  const checklistsPane = document.getElementById("textsChecklistsPane");
  const helpBtn = document.getElementById("textsHelpBtn");
  const isChecklistMode = activeNotebookType === "checklist";

  if (searchWrap) {
    searchWrap.hidden = false;
    searchWrap.setAttribute("aria-hidden", "false");
  }

  syncNotebookSearchUi();

  if (helpBtn) {
    const label = isChecklistMode ? "Checklists help" : "Notes help";
    helpBtn.title = label;
    helpBtn.setAttribute("aria-label", label);
    helpBtn.onclick = () =>
      openExternalUrl(isChecklistMode ? HELP_LINKS.checklists : HELP_LINKS.notes);
  }

  if (notesPane) {
    notesPane.hidden = isChecklistMode;
    notesPane.setAttribute("aria-hidden", String(isChecklistMode));
  }

  if (checklistsPane) {
    checklistsPane.hidden = !isChecklistMode;
    checklistsPane.setAttribute("aria-hidden", String(!isChecklistMode));
  }

  if (isChecklistMode) {
    renderChecklistSearchResults();
  } else {
    renderNoteSearchResults();
  }
}

function renderNoteTabs() {
  ensureActiveNotebookSelection();
  const container = document.getElementById("notesTabsContainer");
  container.innerHTML = "";
  noteTabs.forEach((tabName) => {
    const isActive = activeNotebookType === "note" && tabName === activeNoteTab;
    const btn = el("button", {
      className: `inner-tab-btn ${isActive ? "active" : ""}`,
      type: "button",
      onclick: () => {
        activeNoteTab = tabName;
        activeNotebookType = "note";
        renderNoteTabs();
        renderChecklistTabs();
        renderTextsView();
        renderNoteSearchResults();
      },
    });
    btn.appendChild(
      el("span", {
        className: "notebook-pill-label",
        textContent: tabName,
      })
    );
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
            ensureActiveNotebookSelection(activeNotebookType);
            saveNotes();
            renderNoteTabs();
            renderChecklistTabs();
            renderTextsView();
          }
        }
      },
    });
    btn.appendChild(delIcon);
    container.appendChild(btn);
  });
  const addBtn = el("button", {
    className: "add-tab-btn",
    type: "button",
    textContent: "+",
    title: "Add New Page",
    "aria-label": "Add new page",
    onclick: () => {
      const name = prompt("Enter name for new page:");
      if (name && name.trim()) {
        if (!noteTabs.includes(name.trim())) {
          noteTabs.push(name.trim());
          activeNoteTab = name.trim();
          activeNotebookType = "note";
          saveNotes();
          renderNoteTabs();
          renderChecklistTabs();
          renderTextsView();
        } else toast("Page name already exists.");
      }
    },
  });
  container.appendChild(addBtn);
  updateActiveNoteTextarea();
}

function updateActiveNoteTextarea() {
  const textarea = document.getElementById("notesTextarea");
  const title = document.getElementById("activeNoteTitle");
  if (title) title.textContent = activeNoteTab || "Notes";
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
  const query = String(notesSearchQuery || "").toLowerCase();
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

function createChecklistSearchResultItem(item, numberLookup) {
  const isSubheader = isChecklistSubheader(item);
  const badgeText = isSubheader ? "Section" : String(numberLookup.get(item.id) || 1);
  const itemText = String(item.text || "").trim() || (isSubheader ? "Untitled section" : "Untitled item");

  const button = el("button", {
    className: "checklist-search-result-item",
    type: "button",
    onclick: () => {
      focusChecklistRowInput(item.id, { scrollIntoView: true });
    },
  });

  button.append(
    el("span", {
      className: `checklist-search-result-badge ${isSubheader ? "is-section" : ""}`,
      textContent: badgeText,
    }),
    el("span", {
      className: "checklist-search-result-label",
      textContent: itemText,
    })
  );
  return button;
}

function renderChecklistSearchResults() {
  const resultsContainer = document.getElementById("checklistSearchResults");
  if (!resultsContainer) return;

  resultsContainer.innerHTML = "";
  const checklist = getChecklistById(activeChecklistTabId);
  if (!checklist) return;

  const query = String(checklistSearchQuery || "").trim().toLowerCase();
  if (!query) return;

  const queryWords = query.split(" ").filter(Boolean);
  if (!queryWords.length) return;

  const numberLookup = getChecklistItemNumberLookup(checklist.items);
  const matches = checklist.items.filter((item) => {
    const itemText = String(item.text || "").toLowerCase();
    return queryWords.every((word) => itemText.includes(word));
  });

  if (!matches.length) {
    resultsContainer.innerHTML =
      "<p class='muted tiny checklist-search-empty'>No checklist results found.</p>";
    return;
  }

  matches.forEach((item) => {
    resultsContainer.appendChild(createChecklistSearchResultItem(item, numberLookup));
  });
}

// ===================== CHECKLISTS RENDERING =====================

let activeChecklistTabId = null;
let checklistCreateMenuOpen = false;
let checklistSettingsMenuOpen = false;
const debouncedSaveChecklists = debounce(saveChecklists, 500);

function closeChecklistMenus() {
  const hadOpenMenu =
    checklistCreateMenuOpen || checklistSettingsMenuOpen || Boolean(checklistRowMenuState);
  checklistCreateMenuOpen = false;
  checklistSettingsMenuOpen = false;
  checklistRowMenuState = null;
  return hadOpenMenu;
}

function syncChecklistMenuUi() {
  const createDropdown = document.querySelector(".checklist-create-dropdown");
  const createBtn = createDropdown?.querySelector(".checklist-create-trigger");
  const createMenu = createDropdown?.querySelector(".checklist-action-menu");
  createDropdown?.classList.toggle("open", checklistCreateMenuOpen);
  createBtn?.setAttribute("aria-expanded", String(checklistCreateMenuOpen));
  createMenu?.classList.toggle("open", checklistCreateMenuOpen);

  const settingsDropdown = document.getElementById("checklistSettingsDropdown");
  const settingsBtn = document.getElementById("checklistSettingsBtn");
  const settingsMenu = document.getElementById("checklistSettingsMenu");
  const hasActiveChecklist = Boolean(getChecklistById(activeChecklistTabId));
  const settingsOpen = hasActiveChecklist && checklistSettingsMenuOpen;

  if (settingsDropdown) {
    settingsDropdown.hidden = !hasActiveChecklist;
    settingsDropdown.classList.toggle("open", settingsOpen);
  }
  if (settingsBtn) settingsBtn.setAttribute("aria-expanded", String(settingsOpen));
  if (settingsMenu) settingsMenu.classList.toggle("open", settingsOpen);

  document.querySelectorAll(".checklist-row-menu-dropdown").forEach((dropdown) => {
    const checklistId = String(dropdown.dataset.checklistId || "").trim();
    const itemId = String(dropdown.dataset.itemId || "").trim();
    const isOpen = isChecklistRowMenuOpen(checklistId, itemId);
    const trigger = dropdown.querySelector(".checklist-drag-handle");
    const menu = dropdown.querySelector(".checklist-row-handle-menu");
    dropdown.classList.toggle("open", isOpen);
    if (trigger) trigger.setAttribute("aria-expanded", String(isOpen));
    menu?.classList.toggle("open", isOpen);
  });
}

function handleCreateBlankChecklist() {
  checklistCreateMenuOpen = false;
  const name = prompt("Enter name for new checklist:");
  if (!name || !name.trim()) {
    renderNoteTabs();
    renderChecklistTabs();
    renderTextsView();
    return;
  }
  const newChecklist = createChecklist(name.trim());
  activeChecklistTabId = newChecklist.id;
  activeNotebookType = "checklist";
  renderNoteTabs();
  renderChecklistTabs();
  renderTextsView();
}

function handleCreateChecklistFromTemplate(templateKey) {
  checklistCreateMenuOpen = false;
  const newChecklist = createChecklistFromTemplate(templateKey);
  if (!newChecklist) {
    toast("Template not found.");
    renderNoteTabs();
    renderChecklistTabs();
    renderTextsView();
    return;
  }
  activeChecklistTabId = newChecklist.id;
  activeNotebookType = "checklist";
  renderNoteTabs();
  renderChecklistTabs();
  renderTextsView();
}

function renderChecklistTabs() {
  ensureActiveNotebookSelection();
  const container = document.getElementById("checklistsTabsContainer");
  const actionsContainer = document.getElementById("checklistsSidebarActions");
  if (!container) return;

  container.innerHTML = "";
  if (actionsContainer) {
    actionsContainer.innerHTML = "";
  }

  checklistsDb.checklists.forEach((checklist) => {
    const isActive =
      activeNotebookType === "checklist" && checklist.id === activeChecklistTabId;
    const btn = el("button", {
      className: `inner-tab-btn ${isActive ? "active" : ""}`,
      type: "button",
      onclick: () => {
        closeChecklistMenus();
        activeChecklistTabId = checklist.id;
        activeNotebookType = "checklist";
        renderNoteTabs();
        renderChecklistTabs();
        renderTextsView();
      },
    });
    btn.appendChild(
      el("span", {
        className: "notebook-pill-label",
        textContent: checklist.name,
      })
    );

    // Add remove icon
    const delIcon = el("span", {
      className: "tab-delete-icon",
      textContent: "x",
      title: "Delete Checklist",
      onclick: (e) => {
        e.stopPropagation();
        if (confirm(`Permanently delete checklist "${checklist.name}"?`)) {
          closeChecklistMenus();
          deleteChecklist(checklist.id);
          if (activeChecklistTabId === checklist.id) {
            activeChecklistTabId = checklistsDb.checklists[0]?.id || null;
          }
          ensureActiveNotebookSelection(activeNotebookType);
          renderNoteTabs();
          renderChecklistTabs();
          renderTextsView();
        }
      },
    });
    btn.appendChild(delIcon);

    container.appendChild(btn);
  });

  if (activeChecklistTabId && !getChecklistById(activeChecklistTabId)) {
    activeChecklistTabId = checklistsDb.checklists[0]?.id || null;
  }
  ensureActiveNotebookSelection(activeNotebookType);

  const sidebarActionsTarget = actionsContainer || container;
  const templates = getChecklistTemplatesForCurrentDisciplines();
  const createDropdown = el("div", {
    className: `checklist-action-dropdown checklist-create-dropdown ${
      checklistCreateMenuOpen ? "open" : ""
    }`,
  });
  const createBtn = el("button", {
    className: "add-tab-btn checklist-create-trigger",
    type: "button",
    title: "Add Checklist",
    "aria-label": "Add checklist",
    "aria-haspopup": "menu",
    "aria-expanded": String(checklistCreateMenuOpen),
    onclick: (e) => {
      e.stopPropagation();
      checklistCreateMenuOpen = !checklistCreateMenuOpen;
      if (checklistCreateMenuOpen) checklistSettingsMenuOpen = false;
      syncChecklistMenuUi();
    },
  });
  createBtn.textContent = "+";

  const createMenu = el("div", {
    className: `checklist-action-menu ${checklistCreateMenuOpen ? "open" : ""}`,
    role: "menu",
    "aria-label": "Create checklist",
  });
  createMenu.appendChild(
    el("button", {
      className: "checklist-menu-item",
      type: "button",
      textContent: "Blank checklist",
      onclick: (e) => {
        e.stopPropagation();
        handleCreateBlankChecklist();
      },
    })
  );
  createMenu.appendChild(el("div", { className: "checklist-menu-divider" }));
  createMenu.appendChild(
    el("p", {
      className: "checklist-menu-label",
      textContent: "Start from template",
    })
  );
  if (templates.length) {
    templates.forEach((template) => {
      createMenu.appendChild(
        el("button", {
          className: "checklist-menu-item",
          type: "button",
          textContent: template.name,
          onclick: (e) => {
            e.stopPropagation();
            handleCreateChecklistFromTemplate(template.key);
          },
        })
      );
    });
  } else {
    createMenu.appendChild(
      el("p", {
        className: "checklist-menu-empty",
        textContent: "No templates available for the current discipline selection.",
      })
    );
  }
  createDropdown.append(createBtn, createMenu);
  sidebarActionsTarget.appendChild(createDropdown);

  renderChecklistItems();
}

function renderChecklistItems() {
  const container = document.getElementById("checklistItemsContainer");
  const searchResults = document.getElementById("checklistSearchResults");
  const nameInput = document.getElementById("checklistName");
  const settingsDropdown = document.getElementById("checklistSettingsDropdown");
  const settingsBtn = document.getElementById("checklistSettingsBtn");
  const settingsMenu = document.getElementById("checklistSettingsMenu");
  const resetBtn = document.getElementById("resetChecklistBtn");
  const overrideBtn = document.getElementById("overrideChecklistTemplateBtn");
  const deleteBtn = document.getElementById("deleteChecklistBtn");
  const addItemBtn = document.getElementById("addChecklistItemBtn");
  const addSubheaderBtn = document.getElementById("addChecklistSubheaderBtn");
  if (!container || !nameInput || !addItemBtn || !addSubheaderBtn) return;

  container.innerHTML = "";

  const checklist = getChecklistById(activeChecklistTabId);

  if (!checklist) {
    checklistRowMenuState = null;
    if (searchResults) searchResults.innerHTML = "";
    nameInput.value = "";
    nameInput.disabled = true;
    checklistSettingsMenuOpen = false;
    if (settingsDropdown) settingsDropdown.hidden = true;
    if (settingsBtn) settingsBtn.setAttribute("aria-expanded", "false");
    if (settingsMenu) settingsMenu.classList.remove("open");
    if (resetBtn) resetBtn.hidden = true;
    if (overrideBtn) overrideBtn.hidden = true;
    if (deleteBtn) deleteBtn.hidden = true;
    addItemBtn.disabled = true;
    addSubheaderBtn.disabled = true;
    container.innerHTML = '<p class="muted">Select or create a checklist to get started.</p>';
    return;
  }

  if (
    checklistRowMenuState &&
    (checklistRowMenuState.checklistId !== checklist.id ||
      !checklist.items.some((item) => item.id === checklistRowMenuState.itemId))
  ) {
    checklistRowMenuState = null;
  }

  nameInput.value = checklist.name;
  nameInput.disabled = false;
  if (settingsDropdown) {
    settingsDropdown.hidden = false;
    settingsDropdown.classList.toggle("open", checklistSettingsMenuOpen);
  }
  if (settingsBtn) settingsBtn.setAttribute("aria-expanded", String(checklistSettingsMenuOpen));
  if (settingsMenu) settingsMenu.classList.toggle("open", checklistSettingsMenuOpen);
  if (resetBtn) {
    resetBtn.hidden = !getChecklistTemplateByKey(checklist.templateKey);
  }
  if (overrideBtn) {
    overrideBtn.hidden = !getChecklistTemplateByKey(checklist.templateKey);
  }
  if (deleteBtn) deleteBtn.hidden = false;
  addItemBtn.disabled = false;
  addSubheaderBtn.disabled = false;

  const numberLookup = getChecklistItemNumberLookup(checklist.items);
  checklist.items.forEach((item, index) => {
    const isSubheader = isChecklistSubheader(item);
    const rowMenuOpen = isChecklistRowMenuOpen(checklist.id, item.id);
    const row = el("div", {
      className: `checklist-item-row ${isSubheader ? "checklist-subheader-row" : ""}`,
      "data-item-id": item.id,
    });

    const leadingElement = isSubheader
      ? el("div", {
          className: "checklist-subheader-badge",
          textContent: "Section",
        })
      : el("div", {
          className: "checklist-item-order",
          textContent: String(numberLookup.get(item.id) || 1),
        });

    const input = el("input", {
      type: "text",
      className: `checklist-item-input checklist-row-input ${
        isSubheader ? "checklist-subheader-input" : ""
      }`,
      value: item.text,
      placeholder: isSubheader
        ? "Enter section subheader..."
        : "Enter checklist item...",
      "data-item-id": item.id,
      oninput: (e) => {
        updateChecklistItem(checklist.id, item.id, e.target.value);
      },
    });

    const removeBtn = el("button", {
      className: "checklist-item-remove",
      type: "button",
      title: isSubheader ? "Remove subheader" : "Remove item",
      onclick: () => {
        if (confirm(isSubheader ? "Remove this subheader?" : "Remove this item?")) {
          removeChecklistItem(checklist.id, item.id);
          renderChecklistItems();
        }
      },
    });
    // Use an elegant X for the remove button instead of text
    removeBtn.innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>`;

    const handleMenuDropdown = el("div", {
      className: `checklist-action-dropdown checklist-row-menu-dropdown ${
        rowMenuOpen ? "open" : ""
      }`,
      "data-checklist-id": checklist.id,
      "data-item-id": item.id,
    });

    // Dragging a subheader moves the whole section; dragging an item moves only that row.
    const dragHandle = el("button", {
      className: "checklist-drag-handle",
      type: "button",
      title: isSubheader ? "Drag to move section" : "Drag to move item",
      "aria-label": isSubheader ? "Drag to move section" : "Drag to move item",
      "aria-haspopup": "menu",
      "aria-expanded": String(rowMenuOpen),
    });
    dragHandle.innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="9" cy="12" r="1"></circle><circle cx="9" cy="5" r="1"></circle><circle cx="9" cy="19" r="1"></circle><circle cx="15" cy="12" r="1"></circle><circle cx="15" cy="5" r="1"></circle><circle cx="15" cy="19" r="1"></circle></svg>`;
    dragHandle.draggable = true;
    dragHandle.addEventListener("click", (e) => {
      e.stopPropagation();
      if (Date.now() < checklistHandleMenuSuppressClickUntil) return;
      checklistCreateMenuOpen = false;
      checklistSettingsMenuOpen = false;
      setChecklistRowMenuState(
        checklist.id,
        isChecklistRowMenuOpen(checklist.id, item.id) ? null : item.id
      );
      syncChecklistMenuUi();
    });
    dragHandle.addEventListener("dragstart", (e) => {
      checklistHandleMenuSuppressClickUntil = Date.now() + 250;
      checklistCreateMenuOpen = false;
      checklistSettingsMenuOpen = false;
      checklistRowMenuState = null;
      checklistDragState = {
        checklistId: checklist.id,
        itemId: item.id,
        type: isSubheader ? "section" : "item",
      };
      row.classList.add("checklist-dragging");
      syncChecklistMenuUi();
      if (e.dataTransfer) {
        e.dataTransfer.effectAllowed = "move";
        e.dataTransfer.setData("text/plain", item.id);
      }
    });
    dragHandle.addEventListener("dragend", () => {
      checklistHandleMenuSuppressClickUntil = Date.now() + 250;
      checklistDragState = null;
      row.classList.remove("checklist-dragging");
      clearChecklistDragStyles();
    });

    const rowHandleMenu = el("div", {
      className: `checklist-action-menu checklist-row-handle-menu ${rowMenuOpen ? "open" : ""}`,
      role: "menu",
      "aria-label": "Checklist row actions",
    });
    rowHandleMenu.appendChild(
      el("button", {
        className: "checklist-menu-item",
        type: "button",
        textContent: "Insert item above",
        onclick: (e) => {
          e.stopPropagation();
          checklistRowMenuState = null;
          const newItem = insertChecklistItem(checklist.id, index, "New checklist item");
          renderChecklistItems();
          focusChecklistRowInput(newItem?.id);
        },
      })
    );
    rowHandleMenu.appendChild(
      el("button", {
        className: "checklist-menu-item",
        type: "button",
        textContent: "Insert item below",
        onclick: (e) => {
          e.stopPropagation();
          checklistRowMenuState = null;
          const newItem = insertChecklistItem(checklist.id, index + 1, "New checklist item");
          renderChecklistItems();
          focusChecklistRowInput(newItem?.id);
        },
      })
    );
    rowHandleMenu.appendChild(el("div", { className: "checklist-menu-divider" }));
    rowHandleMenu.appendChild(
      el("button", {
        className: "checklist-menu-item",
        type: "button",
        textContent: "Insert subheader above",
        onclick: (e) => {
          e.stopPropagation();
          checklistRowMenuState = null;
          const newItem = insertChecklistSubheader(checklist.id, index, "New section");
          renderChecklistItems();
          focusChecklistRowInput(newItem?.id);
        },
      })
    );
    rowHandleMenu.appendChild(
      el("button", {
        className: "checklist-menu-item",
        type: "button",
        textContent: "Insert subheader below",
        onclick: (e) => {
          e.stopPropagation();
          checklistRowMenuState = null;
          const newItem = insertChecklistSubheader(checklist.id, index + 1, "New section");
          renderChecklistItems();
          focusChecklistRowInput(newItem?.id);
        },
      })
    );
    handleMenuDropdown.append(dragHandle, rowHandleMenu);

    row.addEventListener("dragover", (e) => {
      if (!checklistDragState || checklistDragState.checklistId !== checklist.id) return;
      e.preventDefault();
      if (e.dataTransfer) e.dataTransfer.dropEffect = "move";
      const rect = row.getBoundingClientRect();
      const before = e.clientY < rect.top + rect.height / 2;
      row.classList.toggle("checklist-drop-before", before);
      row.classList.toggle("checklist-drop-after", !before);
      row.dataset.dropPosition = before ? "before" : "after";
    });

    row.addEventListener("dragleave", (e) => {
      const nextTarget = e.relatedTarget;
      if (nextTarget instanceof Node && row.contains(nextTarget)) return;
      row.classList.remove("checklist-drop-before", "checklist-drop-after");
      delete row.dataset.dropPosition;
    });

    row.addEventListener("drop", (e) => {
      if (!checklistDragState || checklistDragState.checklistId !== checklist.id) return;
      e.preventDefault();
      const before = row.dataset.dropPosition !== "after";
      const moved = moveChecklistRows(
        checklist.id,
        checklistDragState.itemId,
        item.id,
        before
      );
      clearChecklistDragStyles();
      checklistDragState = null;
      if (moved) {
        renderChecklistItems();
      }
    });

    row.append(handleMenuDropdown, leadingElement, input, removeBtn);
    container.appendChild(row);
  });

  renderChecklistSearchResults();
}

function focusChecklistRowInput(itemId, options = {}) {
  if (!itemId) return;
  const { scrollIntoView = false } = options;
  setTimeout(() => {
    const input = document.querySelector(`.checklist-row-input[data-item-id="${itemId}"]`);
    if (input) {
      if (scrollIntoView) {
        input.closest(".checklist-item-row")?.scrollIntoView({
          block: "center",
          behavior: "smooth",
        });
      }
      input.focus();
      input.select();
    }
  }, 50);
}

// Event listeners for checklist tab
document.getElementById("checklistName")?.addEventListener("input", (e) => {
  const checklist = getChecklistById(activeChecklistTabId);
  if (checklist) {
    checklist.name = e.target.value;
    debouncedSaveChecklists();
    renderChecklistTabs();
  }
});

document.getElementById("checklistSettingsBtn")?.addEventListener("click", (e) => {
  e.stopPropagation();
  const checklist = getChecklistById(activeChecklistTabId);
  if (!checklist) return;
  checklistSettingsMenuOpen = !checklistSettingsMenuOpen;
  if (checklistSettingsMenuOpen) checklistCreateMenuOpen = false;
  syncChecklistMenuUi();
});

document.getElementById("resetChecklistBtn")?.addEventListener("click", () => {
  const checklist = getChecklistById(activeChecklistTabId);
  const template = getChecklistTemplateByKey(checklist?.templateKey);
  checklistSettingsMenuOpen = false;
  if (!checklist || !template) {
    syncChecklistMenuUi();
    toast("This checklist does not have template defaults.");
    return;
  }
  if (
    confirm(
      `Reset this checklist to "${template.name}" defaults? This will remove custom items.`
    )
  ) {
    resetChecklistToDefault(activeChecklistTabId);
    renderChecklistTabs();
    return;
  }
  syncChecklistMenuUi();
});

document.getElementById("overrideChecklistTemplateBtn")?.addEventListener("click", () => {
  const checklist = getChecklistById(activeChecklistTabId);
  const template = getChecklistTemplateByKey(checklist?.templateKey);
  checklistSettingsMenuOpen = false;
  if (!checklist || !template) {
    syncChecklistMenuUi();
    toast("This checklist is not linked to a template.");
    return;
  }
  if (
    confirm(
      `Save "${checklist.name}" over the "${template.name}" template? New checklists and reset-to-template actions will use this version.`
    )
  ) {
    overrideChecklistTemplate(checklist.id);
    renderChecklistTabs();
    toast("Template override saved.");
    return;
  }
  syncChecklistMenuUi();
});

document.getElementById("deleteChecklistBtn")?.addEventListener("click", () => {
  const checklist = getChecklistById(activeChecklistTabId);
  checklistSettingsMenuOpen = false;
  if (checklist && confirm(`Delete checklist "${checklist.name}"?`)) {
    deleteChecklist(checklist.id);
    activeChecklistTabId = checklistsDb.checklists[0]?.id || null;
    ensureActiveNotebookSelection(activeNotebookType);
    renderNoteTabs();
    renderChecklistTabs();
    renderTextsView();
    return;
  }
  syncChecklistMenuUi();
});

document.getElementById("addChecklistItemBtn")?.addEventListener("click", () => {
  const checklist = getChecklistById(activeChecklistTabId);
  if (checklist) {
    const newItem = addChecklistItem(checklist.id, "New checklist item");
    renderChecklistItems();
    focusChecklistRowInput(newItem?.id);
  }
});

document.getElementById("addChecklistSubheaderBtn")?.addEventListener("click", () => {
  const checklist = getChecklistById(activeChecklistTabId);
  if (checklist) {
    const newItem = addChecklistSubheader(checklist.id, "New section");
    renderChecklistItems();
    focusChecklistRowInput(newItem?.id);
  }
});

document.addEventListener("click", (e) => {
  if (!checklistCreateMenuOpen && !checklistSettingsMenuOpen && !checklistRowMenuState) return;

  const clickedCreateMenu = e.target.closest(".checklist-create-dropdown");
  const clickedSettingsMenu = e.target.closest("#checklistSettingsDropdown");
  const clickedRowMenu = e.target.closest(".checklist-row-menu-dropdown");

  if (!clickedCreateMenu) checklistCreateMenuOpen = false;
  if (!clickedSettingsMenu) checklistSettingsMenuOpen = false;
  if (!clickedRowMenu) checklistRowMenuState = null;

  syncChecklistMenuUi();
});

document.addEventListener("keydown", (e) => {
  if (e.key !== "Escape") return;
  if (!checklistCreateMenuOpen && !checklistSettingsMenuOpen && !checklistRowMenuState) return;
  closeChecklistMenus();
  syncChecklistMenuUi();
});

window.__aciesAutomation = {
  async getCadAutoSelectTrace(lineLimit = 200) {
    if (!window.pywebview?.api?.get_cad_auto_select_trace) {
      return { status: "error", message: "CAD auto-select trace API is unavailable." };
    }
    return window.pywebview.api.get_cad_auto_select_trace(lineLimit);
  },
  async clearCadAutoSelectTrace() {
    if (!window.pywebview?.api?.clear_cad_auto_select_trace) {
      return { status: "error", message: "CAD auto-select trace API is unavailable." };
    }
    return window.pywebview.api.clear_cad_auto_select_trace();
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

const HIDDEN_PLUGIN_BUNDLES = new Set([
  "ElectricalCommands.GetAttributesCommands.bundle",
]);

function normalizeBundleCoreName(bundleName) {
  return String(bundleName || "")
    .replace("ElectricalCommands.", "")
    .replace(".bundle", "");
}

function isVisiblePluginBundle(bundle) {
  const bundleName = String(bundle?.bundle_name || "");
  if (HIDDEN_PLUGIN_BUNDLES.has(bundleName)) return false;

  const name = String(bundle?.name || "");
  return normalizeBundleCoreName(bundleName || name) !== "GetAttributesCommands";
}

function buildPluginDocUrl(title, pageId) {
  return `https://brainy-seahorse-3c5.notion.site/${title}-${pageId}`;
}

// Keep the desktop app aligned with the current Notion command pages while
// the upstream bundle metadata catches up.
const PLUGIN_DOC_OVERRIDES = {
  AutoLispCommands: {
    replace: true,
    order: ["CLEANUP", "REV", "SUMLENGTHS"],
    commands: {
      CLEANUP:
        "Runs SETBYLAYER on all objects, PURGE ALL, and AUDIT with fixes to clean the current drawing.",
      REV:
        "Draws a revision cloud and inserts a delta triangle with the corresponding revision value.",
      SUMLENGTHS:
        "Calculates the total length of selected lines, arcs, and polylines and reports the result in the command line.",
    },
    links: {
      CLEANUP: buildPluginDocUrl(
        "AutoLispCommands",
        "2b13fdbb662c80479711de7b33eb2737"
      ),
      REV: buildPluginDocUrl(
        "AutoLispCommands",
        "2b13fdbb662c80479711de7b33eb2737"
      ),
      SUMLENGTHS: buildPluginDocUrl(
        "AutoLispCommands",
        "2b13fdbb662c80479711de7b33eb2737"
      ),
    },
  },
  CleanCADCommands: {
    replace: true,
    order: ["EMBEDIMAGES", "EMBEDPDFS", "CLEANTBLK", "CLEANCAD"],
    commands: {
      EMBEDIMAGES:
        "Embeds raster images from XREFs into the drawing by converting them to OLE objects using PowerPoint, preserving orientation.",
      EMBEDPDFS:
        "Embeds PDF underlays into the drawing by converting them to PNG and then OLE objects using PowerPoint.",
      CLEANTBLK:
        "Cleans the title block by exploding blocks, keeping only the title block, detaching XREFs, and embedding images.",
      CLEANCAD:
        "Cleans the entire sheet by embedding XREFs and performing cleanup operations.",
    },
    links: {
      EMBEDIMAGES: buildPluginDocUrl(
        "EMBEDIMAGES",
        "2b03fdbb662c80a08954db8d05ac745f"
      ),
      EMBEDPDFS: buildPluginDocUrl(
        "EMBEDPDFS",
        "2b03fdbb662c806b9716fcd761700898"
      ),
      CLEANTBLK: buildPluginDocUrl(
        "CLEANTBLK",
        "2b03fdbb662c80eb89add09ae9cd99df"
      ),
      CLEANCAD: buildPluginDocUrl(
        "CLEANCAD",
        "2b03fdbb662c805b9df9c80b6ad26900"
      ),
    },
  },
  GeneralCommands: {
    replace: true,
    order: ["OBJMASK", "VPFROMREG", "LAYRED", "LAYGREEN", "AREALABEL"],
    commands: {
      OBJMASK:
        "Creates wipeout objects behind selected text, MText, tables, or polylines to mask the background.",
      VPFROMREG:
        "Creates a viewport in paperspace from a selected region in modelspace, with automatic scaling options.",
      LAYRED:
        "Sets selected layer, region entities, or XREF-related layers to red based on the chosen mode.",
      LAYGREEN: "Sets all layers belonging to a selected XREF to green.",
      AREALABEL:
        "Calculates the area of selected closed polylines and places square-foot labels at their centers.",
    },
    links: {
      OBJMASK: buildPluginDocUrl(
        "OBJMASK",
        "2b13fdbb662c802fb082e4095528ec06"
      ),
      VPFROMREG: buildPluginDocUrl(
        "VPFROMREG",
        "2b13fdbb662c805c963acf926dc03684"
      ),
      LAYRED: buildPluginDocUrl(
        "LAYRED",
        "3e8ae8f68ef04512b3c693c8b94d7503"
      ),
      LAYGREEN: buildPluginDocUrl(
        "LAYGREEN",
        "0401b0c14e7a4581b27371fdc7c635e7"
      ),
      AREALABEL: buildPluginDocUrl(
        "AREALABEL",
        "2b13fdbb662c80dca341f5624feaa488"
      ),
    },
  },
  T24Commands: {
    replace: true,
    order: ["T24", "PDFSHEETS", "IMGSHEETS"],
    commands: {
      T24: "Inserts a T24 PDF form and automatically adds embedded text and images on the final page.",
      PDFSHEETS:
        "Inserts all pages of a selected multi-page PDF as underlays in a grid layout.",
      IMGSHEETS:
        "Inserts all PNG pages from a selected folder into paper space in a grid layout.",
    },
    links: {
      T24: buildPluginDocUrl("T24", "2f53fdbb662c808c881ffd0bd69dc5b1"),
      PDFSHEETS: buildPluginDocUrl(
        "PDFSHEETS",
        "2b13fdbb662c80468573df2164539808"
      ),
      IMGSHEETS: buildPluginDocUrl(
        "IMGSHEETS",
        "67335a8b3ec741e297b5025dd0c413ee"
      ),
    },
  },
  TextCommands: {
    replace: true,
    order: [
      "TEXTCOUNT",
      "TEXTSELECT",
      "TEXTSUM",
      "TEXTSUMEXPORT",
      "TEXTREPLACE",
      "TEXTINCREMENT",
      "TEXTADD",
    ],
    commands: {
      TEXTCOUNT:
        "Counts selected TEXT, MTEXT, and ATTRIB values, groups identical strings, reports totals, and exports a timestamped text report.",
      TEXTSELECT:
        "Filters a mixed selection so only TEXT and MTEXT objects remain selected.",
      TEXTSUM:
        "Sums numeric values parsed from selected TEXT and MTEXT objects.",
      TEXTSUMEXPORT:
        "Finds room-type text near square-footage text and exports grouped totals to T24Output.json.",
      TEXTREPLACE:
        "Replaces the content of selected TEXT and MTEXT objects with a new user-specified value.",
      TEXTINCREMENT:
        "Renumbers selected TEXT and MTEXT objects using a prefix, start/end range, and odd/even filtering.",
      TEXTADD:
        "Adds a specified integer value to numbers found inside selected TEXT and MTEXT objects.",
    },
    links: {
      TEXTCOUNT: buildPluginDocUrl(
        "TEXTCOUNT",
        "fe72b9787d2d4740a82e881a79cb785a"
      ),
      TEXTSELECT: buildPluginDocUrl(
        "TEXTSELECT",
        "5ed6ba4a1b2f4c2090673c79785b77aa"
      ),
      TEXTSUM: buildPluginDocUrl(
        "TEXTSUM",
        "2b13fdbb662c807aa0d6e00948b128d6"
      ),
      TEXTSUMEXPORT: buildPluginDocUrl(
        "TEXTSUMEXPORT",
        "2b13fdbb662c8098967edf32021a5f6c"
      ),
      TEXTREPLACE: buildPluginDocUrl(
        "TEXTREPLACE",
        "2b13fdbb662c80f19897c73b90801e7e"
      ),
      TEXTINCREMENT: buildPluginDocUrl(
        "TEXTINCREMENT",
        "2b13fdbb662c80d8b2f6ef30877488ca"
      ),
      TEXTADD: buildPluginDocUrl(
        "TEXTADD",
        "2b13fdbb662c808c9016dc9a230f7403"
      ),
    },
  },
  LFSCommands: {
    replace: true,
    order: ["LFS"],
    commands: {
      LFS:
        "Inserts the built-in lighting fixture schedule table template and applies local sync data when available.",
    },
    links: {
      LFS: buildPluginDocUrl(
        "LFS",
        "0f3bd29ae7034574b8458136be3e598f"
      ),
    },
  },
};

function applyPluginDocOverrides(coreName, description) {
  const override = PLUGIN_DOC_OVERRIDES[coreName];
  if (!override) return description;

  const baseDescription =
    description && typeof description === "object" ? description : {};
  const commands = {};
  const links = {};
  const renames = override.renames || {};

  if (!override.replace) {
    Object.entries(baseDescription.commands || {}).forEach(([command, summary]) => {
      commands[renames[command] || command] = summary;
    });
    Object.entries(baseDescription.links || {}).forEach(([command, link]) => {
      links[renames[command] || command] = link;
    });
  }

  Object.assign(commands, override.commands || {});
  Object.assign(links, override.links || {});

  if (Object.keys(commands).length === 0 && Object.keys(links).length === 0) {
    return description;
  }

  const order = Array.isArray(override.order) ? override.order : [];
  if (order.length === 0) {
    return {
      ...baseDescription,
      commands,
      links,
    };
  }

  const orderedCommands = {};
  const orderedLinks = {};

  order.forEach((command) => {
    if (!(command in commands)) return;
    orderedCommands[command] = commands[command];
    if (links[command]) orderedLinks[command] = links[command];
  });

  Object.keys(commands).forEach((command) => {
    if (command in orderedCommands) return;
    orderedCommands[command] = commands[command];
    if (links[command]) orderedLinks[command] = links[command];
  });

  return {
    ...baseDescription,
    commands: orderedCommands,
    links: orderedLinks,
  };
}

// Fetch description from specific GitHub RAW url based on bundle name
async function fetchDescriptionForBundle(bundleName) {
  // Remove prefix/suffix to get core name. e.g., "ElectricalCommands.CleanCADCommands.bundle" -> "CleanCADCommands"
  const coreName = normalizeBundleCoreName(bundleName);

  // Check Cache
  if (DESCRIPTION_CACHE[coreName]) return DESCRIPTION_CACHE[coreName];

  // Construct RAW URL. Pattern: AutoCADCommands/<CoreName>/<CoreName>_descriptions.json
  const url = `https://raw.githubusercontent.com/jacobhusband/ElectricalCommands/main/AutoCADCommands/${coreName}/${coreName}_descriptions.json`;

  try {
    const res = await fetch(url);
    if (!res.ok) throw new Error("Not found");
    const json = await res.json();
    const resolvedDescription = applyPluginDocOverrides(coreName, json);
    DESCRIPTION_CACHE[coreName] = resolvedDescription; // Cache it
    return resolvedDescription;
  } catch (e) {
    console.warn(`Could not fetch description for ${coreName}`, e);
    const fallbackDescription = applyPluginDocOverrides(coreName, null);
    if (fallbackDescription) {
      DESCRIPTION_CACHE[coreName] = fallbackDescription;
      return fallbackDescription;
    }
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
    const data = Array.isArray(response.data)
      ? response.data.filter(isVisiblePluginBundle)
      : [];
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
  const visibleBundles = Array.isArray(bundles)
    ? bundles.filter(isVisiblePluginBundle)
    : [];

  container.innerHTML = "";
  if (visibleBundles.length === 0) {
    container.textContent = "No command bundles found.";
    return;
  }

  const tagJobs = [];

  // Process each bundle
  for (const bundle of visibleBundles) {
    // Normalize name
    const coreName = normalizeBundleCoreName(bundle.name);
    const detailsId = `bundle-details-${String(
      bundle.bundle_name || coreName || "bundle"
    ).replace(/[^a-z0-9_-]+/gi, "-")}`;

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
    } else if (bundle.state === "not_published") {
      statusClass = "not-installed";
      statusTitle = "Release asset not published yet";
      btnText = "Unavailable";
      btnClass = "";
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

    const body = el("div", {
      className: "release-card-body",
      id: detailsId,
      hidden: true,
    });
    const detailsState = el("div", {
      className: "release-card-details-state",
      textContent: "Loading details...",
    });
    const tags = el("div", { className: "command-tags" });
    body.append(detailsState, tags);

    const footer = el("div", { className: "release-card-footer" });
    const detailsBtn = el("button", {
      type: "button",
      className: "release-card-toggle",
      textContent: "Details",
      "aria-controls": detailsId,
      "aria-expanded": "false",
    });
    const btn = el("button", {
      className: `btn ${btnClass}`.trim(),
      textContent: btnText,
    });

    const setDetailsExpanded = (expanded) => {
      body.hidden = !expanded;
      card.classList.toggle("details-expanded", expanded);
      detailsBtn.textContent = expanded ? "Hide details" : "Details";
      detailsBtn.setAttribute("aria-expanded", String(expanded));
    };

    detailsBtn.addEventListener("click", () => {
      setDetailsExpanded(body.hidden);
    });

    btn.dataset.bundleName = bundle.bundle_name;
    if (bundle.state === "not_published") {
      btn.disabled = true;
      btn.title = "This bundle is not yet available as a published release asset.";
    } else {
      btn.dataset.actionType = btnText;
    }

    // Note: We rely on 'bundle.asset' from the backend for installation URL.
    // If backend doesn't provide it, we assume standard release naming.
    if (bundle.state !== "installed" && bundle.asset) {
      btn.dataset.asset = JSON.stringify(bundle.asset);
    }

    footer.append(detailsBtn, btn);
    card.append(header, body, footer);
    container.append(card);

    const descriptionPromise = fetchDescriptionForBundle(bundle.name)
      .then((description) => {
        const commands = description?.commands
          ? Object.keys(description.commands)
          : [];

        if (commands.length > 0) {
          detailsState.textContent =
            commands.length === 1 ? "1 command" : `${commands.length} commands`;

          commands.forEach((cmd) => {
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
        } else {
          detailsState.textContent = "No extra details available.";
        }
      })
      .catch(() => {
        detailsState.textContent = "No extra details available.";
      });
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

function isToolsTabActive() {
  const panel = document.getElementById("tools-panel");
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
      if (isToolsTabActive()) {
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

  const actionType = String(button.dataset.actionType || "").trim();
  const bundleName = String(button.dataset.bundleName || "").trim();
  const bundleLabel = normalizeBundleCoreName(bundleName) || "AutoCAD Plugin";
  const actionVerb =
    actionType === "Uninstall"
      ? "Uninstalling"
      : actionType === "Update"
        ? "Updating"
        : "Installing";
  const activityId = beginActivity({
    activityId: createActivityId(`bundle_${bundleName || actionType}`),
    label: bundleLabel,
    message: `${actionVerb} plugin...`,
    progress: 25,
    kind: "plugin",
    openFolderLabel:
      actionType === "Uninstall" ? "Open Plugins Folder" : "Open Folder",
  });
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
    completeActivity(activityId, {
      message: `${actionType} successful.`,
      openFolderPath: String(
        response.bundlePath || response.pluginsFolderPath || ""
      ).trim(),
      openFolderLabel:
        actionType === "Uninstall" ? "Open Plugins Folder" : "Open Folder",
    });
  } catch (err) {
    failActivity(activityId, {
      message: err?.message || `${actionType} failed.`,
    });
    return;
    toast(`⚠️ ${err.message}`, 5000);
  } finally {
    await ensureBundlesRendered({ force: true });

  }
}

// ===================== TOOLS & SCRIPTS =====================
window.updateToolStatus = function (toolId, message) {
  const abortBtn = document.getElementById("abortBtn");
  const nextMessage = String(message || "").trim();
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

  if (!nextMessage) {
    if (toolId) {
      const activityId = activeToolActivityIds.get(String(toolId || "").trim());
      if (activityId) {
        updateActivity(activityId, {
          message: getActivityById(activityId)?.message || "Working...",
        });
      }
    }
    return;
  }

  initActivityTray();
  updateActivityStatusFromPayload({
    toolId,
    activityId: activeToolActivityIds.get(String(toolId || "").trim()) || "",
    message: nextMessage,
  });
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
  resolveCadLaunchContextForTool();
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
  launchContext: null,
  running: false,
  activeJobId: "",
  activeActivityId: "",
  panelSchedulePollTimer: 0,
  lastHandledTerminalJobId: "",
  lastPanelScheduleStatus: null,
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
    breakerPaths: [],
    directoryPaths: [],
    breakerFiles: [],
    directoryFiles: [],
  };
}

function normalizeCircuitBreakerPaths(paths) {
  if (!Array.isArray(paths)) {
    paths = paths ? [paths] : [];
  }
  return paths
    .map((path) => String(path || "").trim())
    .filter(Boolean);
}

function normalizeCircuitBreakerFiles(files) {
  if (!files) return [];
  if (Array.isArray(files)) return files.filter(Boolean);
  if (typeof files.length === "number") {
    return Array.from(files).filter(Boolean);
  }
  return [files].filter(Boolean);
}

function hasCircuitBreakerPhotoSelection(paths = [], files = []) {
  return (
    normalizeCircuitBreakerPaths(paths).length > 0 ||
    normalizeCircuitBreakerFiles(files).length > 0
  );
}

function getCircuitBreakerPhotoNames(paths = [], files = []) {
  const names = normalizeCircuitBreakerPaths(paths).map((path) => {
    const parts = String(path).split(/[\\/]/);
    return parts[parts.length - 1] || path;
  });
  normalizeCircuitBreakerFiles(files).forEach((file) => {
    names.push(file.name || "upload");
  });
  return names;
}

function getCircuitBreakerPhotoSummary(paths = [], files = []) {
  const names = getCircuitBreakerPhotoNames(paths, files);
  if (!names.length) {
    return {
      label: "No photos selected",
      title: "",
    };
  }
  if (names.length === 1) {
    return {
      label: names[0],
      title: names[0],
    };
  }
  return {
    label: `${names.length} photos selected`,
    title: names.join("\n"),
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
    hasCircuitBreakerPhotoSelection(panel.breakerPaths, panel.breakerFiles) &&
    hasCircuitBreakerPhotoSelection(panel.directoryPaths, panel.directoryFiles)
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

function getCircuitBreakerLaunchDefaultDirectory() {
  return getLaunchContextProjectRoot(circuitBreakerState.launchContext);
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

function setCircuitBreakerStatus(message, isError = false) {
  const statusEl = document.getElementById("cbRunStatus");
  if (!statusEl) return;
  statusEl.textContent = message || "";
  statusEl.classList.toggle("text-danger", Boolean(isError));
}

function getPanelScheduleToolCard() {
  return document.getElementById("toolCircuitBreaker");
}

function setPanelScheduleToolCardStatus(message, { running = false, error = false } = {}) {
  const text = String(message || "").trim();
  if (text || error) {
    window.updateToolStatus(
      "toolCircuitBreaker",
      error ? `ERROR: ${text || "Panel Schedule AI failed."}` : text
    );
  }
  const card = getPanelScheduleToolCard();
  if (!card || error) return;
  card.classList.toggle("running", Boolean(running));
}

function clearPanelScheduleStatusPoll() {
  if (circuitBreakerState.panelSchedulePollTimer) {
    clearTimeout(circuitBreakerState.panelSchedulePollTimer);
    circuitBreakerState.panelSchedulePollTimer = 0;
  }
}

function schedulePanelScheduleStatusPoll(jobId, delay = 1000) {
  clearPanelScheduleStatusPoll();
  if (!jobId || circuitBreakerState.activeJobId !== jobId) return;
  circuitBreakerState.panelSchedulePollTimer = window.setTimeout(() => {
    void pollPanelScheduleBackgroundStatus(jobId);
  }, Math.max(0, Number(delay) || 0));
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
    const summary = getCircuitBreakerPhotoSummary(
      activePanel?.breakerPaths,
      activePanel?.breakerFiles
    );
    breakerFile.textContent = summary.label;
    breakerFile.title = summary.title || summary.label;
    breakerFile.dataset.empty =
      activePanel &&
      hasCircuitBreakerPhotoSelection(
        activePanel.breakerPaths,
        activePanel.breakerFiles
      )
        ? "false"
        : "true";
  }
  if (directoryFile) {
    const summary = getCircuitBreakerPhotoSummary(
      activePanel?.directoryPaths,
      activePanel?.directoryFiles
    );
    directoryFile.textContent = summary.label;
    directoryFile.title = summary.title || summary.label;
    directoryFile.dataset.empty =
      activePanel &&
      hasCircuitBreakerPhotoSelection(
        activePanel.directoryPaths,
        activePanel.directoryFiles
      )
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
      activePanel &&
      hasCircuitBreakerPhotoSelection(
        activePanel.breakerPaths,
        activePanel.breakerFiles
      )
        ? "true"
        : "false";
    breakerDrop.classList.toggle("is-disabled", circuitBreakerState.running);
  }
  if (directoryDrop) {
    directoryDrop.dataset.hasFile =
      activePanel &&
      hasCircuitBreakerPhotoSelection(
        activePanel.directoryPaths,
        activePanel.directoryFiles
      )
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
      `${counts.ready}/${counts.total} panel${counts.total === 1 ? "" : "s"} ready. Each panel needs at least one breaker photo and at least one directory photo.`
    );
  }
}

function resetCircuitBreakerForm({ launchContext = null } = {}) {
  clearPanelScheduleStatusPoll();
  circuitBreakerState.panels = [];
  circuitBreakerState.activePanelId = "";
  circuitBreakerState.nextPanelNumber = 1;
  circuitBreakerState.outputMode = "new";
  circuitBreakerState.newOutputExtension = "xlsx";
  circuitBreakerState.newOutputPath = "";
  circuitBreakerState.existingOutputPath = "";
  circuitBreakerState.launchContext = deepCloneJson(launchContext, null);
  circuitBreakerState.running = false;
  circuitBreakerState.activeJobId = "";
  circuitBreakerState.activeActivityId = "";
  circuitBreakerState.lastHandledTerminalJobId = "";
  circuitBreakerState.lastPanelScheduleStatus = null;
  ensureCircuitBreakerPanels();
  updateCircuitBreakerUi();
}

function setCircuitBreakerFiles(kind, files) {
  const panel = getActiveCircuitBreakerPanel();
  if (!panel) return;
  const nextFiles = normalizeCircuitBreakerFiles(files);
  if (kind === "breaker") {
    panel.breakerFiles = nextFiles;
    panel.breakerPaths = [];
  } else {
    panel.directoryFiles = nextFiles;
    panel.directoryPaths = [];
  }
  updateCircuitBreakerUi();
}

function setCircuitBreakerPaths(kind, paths) {
  const panel = getActiveCircuitBreakerPanel();
  if (!panel) return;
  const nextPaths = normalizeCircuitBreakerPaths(paths);
  if (kind === "breaker") {
    panel.breakerPaths = nextPaths;
    panel.breakerFiles = [];
  } else {
    panel.directoryPaths = nextPaths;
    panel.directoryFiles = [];
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
  const defaultDirectory = getCircuitBreakerLaunchDefaultDirectory();
  try {
    const result = await window.pywebview.api.select_files({
      allow_multiple: true,
      file_types: ["Image Files (*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tif;*.tiff;*.heic;*.heif)"],
      default_directory: defaultDirectory,
    });
    if (result.status === "success" && result.paths?.length) {
      setCircuitBreakerPaths(kind, result.paths);
    }
  } catch (e) {
    toast("Error selecting photos.");
  }
}

async function selectCircuitBreakerSchedulePath(mode) {
  if (!window.pywebview?.api) {
    toast("File picker is unavailable.");
    return;
  }
  const defaultDirectory = getCircuitBreakerLaunchDefaultDirectory();
  if (mode === "new") {
    try {
      const outputExtension = normalizeCircuitBreakerOutputExtension(
        circuitBreakerState.newOutputExtension
      );
      const selection = await window.pywebview.api.select_template_save_location(
        defaultDirectory || null,
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
      default_directory: defaultDirectory,
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
  const launchContext = resolveCadLaunchContextForTool();
  const dlg = document.getElementById("circuitBreakerDlg");
  if (!dlg) return;
  if (!circuitBreakerState.running) {
    resetCircuitBreakerForm({ launchContext });
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
  input.multiple = true;
  input.onchange = () => {
    const files = Array.from(input.files || []);
    if (files.length) setCircuitBreakerFiles(kind, files);
  };
  input.click();
}

function handleCircuitBreakerDrop(kind, e, zone) {
  e.preventDefault();
  if (zone) zone.classList.remove("is-dragover");
  if (circuitBreakerState.running) return;
  const files = Array.from(e.dataTransfer?.files || []);
  if (files.length) setCircuitBreakerFiles(kind, files);
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

async function filesToUploadPayloads(files) {
  const payloads = await Promise.all(
    normalizeCircuitBreakerFiles(files).map((file) => fileToUploadPayload(file))
  );
  return payloads.filter(Boolean);
}

async function pollPanelScheduleBackgroundStatus(jobId) {
  if (!jobId || circuitBreakerState.activeJobId !== jobId) return;
  if (!window.pywebview?.api?.get_panel_schedule_background_status) return;
  try {
    const payload = await window.pywebview.api.get_panel_schedule_background_status(
      jobId
    );
    await handlePanelScheduleBackgroundUpdate(payload, { source: "poll" });
  } catch (e) {
    if (circuitBreakerState.activeJobId === jobId) {
      schedulePanelScheduleStatusPoll(jobId, 1000);
    }
  }
}

async function handlePanelScheduleBackgroundUpdate(payload, { source = "push" } = {}) {
  if (!payload) return;
  const jobId = String(
    payload.jobId || circuitBreakerState.activeJobId || ""
  ).trim();
  if (!jobId) return;

  const status = String(payload.status || "").trim().toLowerCase();
  const activityId = String(
    payload.activityId || circuitBreakerState.activeActivityId || ""
  ).trim();
  const parsedPanelCount = Number(payload.panelCount);
  const panelCount = Number.isFinite(parsedPanelCount)
    ? Math.max(1, parsedPanelCount)
    : 1;
  const parsedSuccess = Number(payload.successCount);
  const parsedFailure = Number(payload.failureCount);
  const successCount = Number.isFinite(parsedSuccess)
    ? Math.max(0, parsedSuccess)
    : status === "success"
      ? 1
      : 0;
  const failureCount = Number.isFinite(parsedFailure)
    ? Math.max(0, parsedFailure)
    : status === "success"
      ? 0
      : 1;
  const completedCount = Math.min(panelCount, successCount + failureCount);
  const results = Array.isArray(payload.results) ? payload.results : [];
  const openFolderPath = String(
    payload.outputFolder ||
      (payload.outputPath ? payload.outputPath.split(/[\\/]/).slice(0, -1).join("\\") : "")
  ).trim();
  circuitBreakerState.lastPanelScheduleStatus = {
    jobId,
    status,
    source,
    activityId,
  };
  if (activityId) {
    circuitBreakerState.activeActivityId = activityId;
  }

  if (status === "running") {
    window.updateActivityStatus({
      ...payload,
      toolId: "toolCircuitBreaker",
      activityId: activityId || getActivityIdForTool("toolCircuitBreaker", { create: true }),
      label: "Panel Schedule AI",
      message:
        payload.message ||
        `Running ${panelCount} panel${panelCount === 1 ? "" : "s"} in background...`,
      panelCount,
      completedCount,
      openFolderPath,
    });
    if (jobId === circuitBreakerState.activeJobId) {
      schedulePanelScheduleStatusPoll(jobId, 1000);
    }
    return;
  }

  if (status === "not_found") {
    if (jobId === circuitBreakerState.activeJobId) {
      schedulePanelScheduleStatusPoll(jobId, 1000);
    }
    return;
  }

  if (status !== "success" && status !== "error") return;
  if (circuitBreakerState.lastHandledTerminalJobId === jobId) return;
  if (circuitBreakerState.activeJobId && jobId !== circuitBreakerState.activeJobId) {
    return;
  }

  circuitBreakerState.lastHandledTerminalJobId = jobId;
  circuitBreakerState.activeJobId = "";
  const terminalActivityId =
    activityId || getActivityIdForTool("toolCircuitBreaker", { create: true });
  circuitBreakerState.activeActivityId = "";
  clearPanelScheduleStatusPoll();
  circuitBreakerState.running = false;
  updateCircuitBreakerUi();

  if (status === "success") {
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

    window.updateActivityStatus({
      ...payload,
      toolId: "toolCircuitBreaker",
      activityId: terminalActivityId,
      label: "Panel Schedule AI",
      status:
        failureCount > 0 ? ACTIVITY_STATUS.WARNING : ACTIVITY_STATUS.SUCCESS,
      message,
      panelCount,
      completedCount,
      openFolderPath,
      openFolderLabel: "Open Folder",
    });
    return;
  }

  const baseMessage = payload.message || "Panel Schedule AI failed.";
  const failedNames = results
    .filter((item) => item?.status === "error")
    .map((item) => item?.panelName || item?.panelId)
    .filter(Boolean);
  const listed = failedNames.length
    ? ` Failed: ${failedNames.slice(0, 2).join(", ")}${failedNames.length > 2 ? ", ..." : ""}.`
    : "";
  const statusMessage = `${baseMessage}${listed}`;
  window.updateActivityStatus({
    ...payload,
    toolId: "toolCircuitBreaker",
    activityId: terminalActivityId,
    label: "Panel Schedule AI",
    status: ACTIVITY_STATUS.ERROR,
    message: statusMessage,
    panelCount,
    completedCount,
    openFolderPath,
    openFolderLabel: "Open Folder",
  });
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
    const breakerUploads = await filesToUploadPayloads(panel.breakerFiles);
    const directoryUploads = await filesToUploadPayloads(panel.directoryFiles);
    panels.push({
      panelId: panel.id,
      panelName: panel.panelName?.trim() || panel.label || "PANEL",
      breakerPaths: [...normalizeCircuitBreakerPaths(panel.breakerPaths)],
      directoryPaths: [...normalizeCircuitBreakerPaths(panel.directoryPaths)],
      breakerUploads,
      directoryUploads,
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
    breakerPaths: firstPanel.breakerPaths || [],
    directoryPaths: firstPanel.directoryPaths || [],
    breakerUploads: firstPanel.breakerUploads || [],
    directoryUploads: firstPanel.directoryUploads || [],
    panelName: firstPanel.panelName || "",
  };
  const activityId = beginActivity({
    activityId: createActivityId("toolCircuitBreaker"),
    toolId: "toolCircuitBreaker",
    label: "Panel Schedule AI",
    message: "Preparing panel schedule job...",
    progress: 10,
  });
  payload.activityId = activityId;

  try {
    const res = await window.pywebview.api.run_panel_schedule_background(payload);
    if (res?.status === "error") {
      clearPanelScheduleStatusPoll();
      circuitBreakerState.activeJobId = "";
      circuitBreakerState.activeActivityId = "";
      circuitBreakerState.running = false;
      updateCircuitBreakerUi();
      failActivity(activityId, {
        message: res.message || "Failed to start Panel Schedule AI.",
      });
      return;
    }

    const jobId = String(res?.jobId || "").trim();
    if (!jobId) {
      clearPanelScheduleStatusPoll();
      circuitBreakerState.activeJobId = "";
      circuitBreakerState.activeActivityId = "";
      circuitBreakerState.running = false;
      updateCircuitBreakerUi();
      failActivity(activityId, {
        message: "Panel Schedule AI did not return a job id.",
      });
      return;
    }

    const parsedPanelCount = Number(res?.panelCount);
    const jobPanelCount = Number.isFinite(parsedPanelCount)
      ? Math.max(1, parsedPanelCount)
      : panels.length;
    const runningMessage = `Running ${jobPanelCount} panel${jobPanelCount === 1 ? "" : "s"} in background...`;
    circuitBreakerState.activeJobId = jobId;
    circuitBreakerState.activeActivityId = String(
      res?.activityId || activityId
    ).trim() || activityId;
    circuitBreakerState.lastHandledTerminalJobId = "";
    circuitBreakerState.lastPanelScheduleStatus = {
      jobId,
      status: "running",
      source: "start",
      activityId: circuitBreakerState.activeActivityId,
    };
    updateActivity(circuitBreakerState.activeActivityId, {
      message: runningMessage,
      progress: 20,
      panelCount: jobPanelCount,
      completedCount: 0,
      openFolderPath:
        circuitBreakerState.outputMode === "existing"
          ? outputPath.split(/[\\/]/).slice(0, -1).join("\\")
          : outputPath.split(/[\\/]/).slice(0, -1).join("\\"),
    });
    schedulePanelScheduleStatusPoll(jobId, 1000);
    closeCircuitBreaker();
  } catch (e) {
    clearPanelScheduleStatusPoll();
    circuitBreakerState.activeJobId = "";
    circuitBreakerState.activeActivityId = "";
    circuitBreakerState.running = false;
    updateCircuitBreakerUi();
    failActivity(activityId, {
      message: e?.message || "Failed to start Panel Schedule AI.",
    });
  }
}

window.handlePanelScheduleResult = async function (payload) {
  await handlePanelScheduleBackgroundUpdate(payload, { source: "push" });
};

const debouncedSaveLightingSchedule = debounce(
  () => persistActiveLightingSchedule({ reason: "edit" }),
  500
);
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

function extractLightingScheduleProjectIdFromPath(rawPath) {
  const path = normalizeLightingSchedulePath(rawPath);
  if (!path) return "";
  const segments = path.split("\\").filter(Boolean);
  for (let i = segments.length - 1; i >= 0; i -= 1) {
    const match = segments[i].trim().match(LIGHTING_PROJECT_SEGMENT_REGEX);
    if (match) return String(match[1] || "").trim();
  }
  return "";
}

function getLightingScheduleProjectId(project, schedule = null) {
  const explicitId = String(project?.id || "").trim();
  if (explicitId) return explicitId;

  const pathCandidates = [
    project?.path,
    project?.localProjectPath,
    project?.workroomRootPath,
    schedule?.targetDwgPath,
  ];
  for (const candidate of pathCandidates) {
    const extracted = extractLightingScheduleProjectIdFromPath(candidate);
    if (extracted) return extracted;
  }

  const displayName = String(project?.name || project?.nick || "").trim();
  return displayName ? `name:${displayName.toLowerCase()}` : "";
}

function getLightingScheduleLoadedVersion(schedule) {
  const value = Number(schedule?._storeVersion);
  return Number.isFinite(value) && value > 0 ? value : 0;
}

function applyLightingScheduleRecordToProject(project, record) {
  if (!project || !record?.schedule) return null;
  const schedule = ensureLightingSchedule(project).schedule;
  const canonical = buildCanonicalLightingSchedule(record.schedule);
  schedule.rows = canonical.rows.map((row) => createLightingScheduleRow(row));
  schedule.generalNotes = canonical.generalNotes;
  schedule.notes = canonical.notes;
  schedule.targetDwgPath = normalizeLightingSchedulePath(record.targetDwgPath);
  schedule._storeVersion = Number(record.version || 0);
  schedule._storeUpdatedAtUtc = String(record.updatedAtUtc || "").trim();
  schedule._storeUpdatedBy = String(record.updatedBy || "").trim();
  schedule._tableHandle = String(record.tableHandle || "").trim();
  return schedule;
}

async function loadLightingScheduleFromCentralStore(
  project,
  { quiet = false, preserveStatus = false } = {}
) {
  const schedule = project ? ensureLightingSchedule(project).schedule : null;
  if (!project || !schedule) return null;
  if (!window.pywebview?.api?.get_lighting_schedule_record) return schedule;

  const projectId = getLightingScheduleProjectId(project, schedule);
  if (!projectId) {
    if (!quiet) {
      setLightingScheduleSyncStatus(
        "Lighting schedule sync needs a project ID or a project-number path before it can use the central store."
      );
    }
    return schedule;
  }

  try {
    const response = await window.pywebview.api.get_lighting_schedule_record(
      projectId
    );
    if (response?.status !== "success") {
      throw new Error(response?.message || "Could not load lighting schedule.");
    }
    if (!response.exists || !response.data) {
      schedule._storeVersion = 0;
      schedule._storeUpdatedAtUtc = "";
      schedule._storeUpdatedBy = "";
      if (!preserveStatus && !quiet) {
        setLightingScheduleSyncStatus(
          "Central store is ready. The first save from desktop or AutoCAD will create the linked schedule record."
        );
      }
      renderLightingSchedule(schedule);
      return schedule;
    }

    const applied = applyLightingScheduleRecordToProject(project, response.data);
    if (applied) {
      lightingScheduleSyncDirty = false;
      renderLightingSchedule(applied);
      if (!preserveStatus && !quiet) {
        setLightingScheduleSyncStatus(
          `Loaded central schedule version ${response.data.version} for ${projectId}.`
        );
      }
    }
    return applied;
  } catch (error) {
    reportClientError("Failed to load lighting schedule from central store", error);
    if (!quiet) {
      setLightingScheduleSyncStatus(
        `Central store load failed: ${error?.message || "Unknown error"}.`
      );
    }
    return schedule;
  }
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
    return "Select a project to load its centralized lighting fixture schedule.";
  }
  const projectId = getLightingScheduleProjectId(project, schedule);
  if (!projectId) {
    return "Add a project ID or use a project-number path so the lighting schedule can use the shared central store.";
  }
  const targetDwgPath = normalizeLightingSchedulePath(schedule.targetDwgPath);
  const segments = [];
  segments.push(`Central key: ${projectId}`);
  if (getLightingScheduleLoadedVersion(schedule) > 0) {
    segments.push(`Version: ${getLightingScheduleLoadedVersion(schedule)}`);
  }
  if (schedule._storeUpdatedAtUtc) {
    segments.push(
      `Last update: ${formatLightingScheduleSyncTimestamp(
        schedule._storeUpdatedAtUtc
      )}`
    );
  }
  if (schedule._storeUpdatedBy) {
    segments.push(`Updated by: ${schedule._storeUpdatedBy}`);
  }
  if (targetDwgPath) {
    segments.push(`Target DWG: ${targetDwgPath}`);
  } else {
    segments.push("Choose Target DWG to bind a drawing for AutoCAD table updates.");
  }
  segments.push("Run LFSOPEN in AutoCAD to edit the same schedule inside AutoCAD.");
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
  if (pullBtn) {
    pullBtn.disabled = !hasProject || lightingScheduleSyncSaving;
    pullBtn.textContent = "Reload";
    pullBtn.title = "Reload the schedule from the central store";
  }
  if (pushBtn) {
    pushBtn.disabled = !hasProject || lightingScheduleSyncSaving;
    pushBtn.textContent = "Save Now";
    pushBtn.title = "Write the current schedule to the central store immediately";
  }

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

function stopLightingScheduleSyncWatcher() {
  if (lightingScheduleSyncPollTimer) {
    clearInterval(lightingScheduleSyncPollTimer);
    lightingScheduleSyncPollTimer = null;
  }
}

function startLightingScheduleSyncWatcher() {
  stopLightingScheduleSyncWatcher();
  if (!window.pywebview?.api?.get_lighting_schedule_version) return;
  lightingScheduleSyncPollTimer = setInterval(() => {
    pollLightingScheduleCentralStore();
  }, 2500);
}

async function pollLightingScheduleCentralStore() {
  const dlg = document.getElementById("lightingScheduleDlg");
  if (!dlg?.open || lightingScheduleSyncDirty || lightingScheduleSyncSaving) {
    return;
  }

  const project = getActiveLightingScheduleProject();
  const schedule = getActiveLightingSchedule();
  if (!project || !schedule) return;

  const projectId = getLightingScheduleProjectId(project, schedule);
  if (!projectId) return;

  try {
    const response = await window.pywebview.api.get_lighting_schedule_version(
      projectId
    );
    if (response?.status !== "success") return;
    const versionInfo = response.data || {};
    const remoteVersion = Number(versionInfo.version || 0);
    if (
      !versionInfo.exists ||
      remoteVersion <= getLightingScheduleLoadedVersion(schedule)
    ) {
      return;
    }

    await loadLightingScheduleFromCentralStore(project, {
      quiet: true,
      preserveStatus: true,
    });
    setLightingScheduleSyncStatus(
      `Detected external lighting schedule update. Reloaded central version ${remoteVersion}.`
    );
  } catch (error) {
    console.warn("Lighting schedule central-store poll failed:", error);
  }
}

async function persistActiveLightingSchedule({
  reason = "save",
  showStatus = false,
  showToast = false,
} = {}) {
  const project = getActiveLightingScheduleProject();
  const schedule = getActiveLightingSchedule();
  if (!project || !schedule) return null;
  if (!window.pywebview?.api?.save_lighting_schedule_record) return schedule;

  const projectId = getLightingScheduleProjectId(project, schedule);
  if (!projectId) {
    const message =
      "Lighting schedule sync needs a project ID or a project-number path before it can save to the central store.";
    setLightingScheduleSyncStatus(message);
    if (showToast) toast(message);
    return null;
  }

  try {
    lightingScheduleSyncSaving = true;
    renderLightingScheduleSyncControls(project, schedule);
    const response = await window.pywebview.api.save_lighting_schedule_record(
      projectId,
      {
        schedule: buildCanonicalLightingSchedule(schedule),
        targetDwgPath: normalizeLightingSchedulePath(schedule.targetDwgPath),
        tableHandle: String(schedule._tableHandle || "").trim(),
        expectedVersion: getLightingScheduleLoadedVersion(schedule),
        updatedBy: "desktop",
      }
    );
    if (response?.status !== "success" || !response.data) {
      throw new Error(response?.message || "Could not save lighting schedule.");
    }

    applyLightingScheduleRecordToProject(project, response.data);
    lightingScheduleSyncDirty = false;
    renderLightingSchedule(schedule);

    if (showStatus) {
      const conflictMessage = response.data.conflict
        ? ` Central version ${response.data.previousVersion} was overwritten by this ${reason}.`
        : "";
      setLightingScheduleSyncStatus(
        `Saved lighting schedule to the central store as version ${response.data.version}.${conflictMessage}`
      );
    }
    if (showToast) {
      toast("Lighting schedule saved.");
    }
    return response.data;
  } catch (error) {
    reportClientError("Failed to save lighting schedule to central store", error);
    if (showStatus || showToast) {
      setLightingScheduleSyncStatus(
        `Lighting schedule save failed: ${error?.message || "Unknown error"}.`
      );
    }
    if (showToast) {
      toast("Lighting schedule save failed.");
    }
    return null;
  } finally {
    lightingScheduleSyncSaving = false;
    renderLightingScheduleSyncControls(project, schedule);
  }
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
    await persistActiveLightingSchedule({
      reason: "target DWG update",
      showStatus: true,
    });
  } catch (error) {
    reportClientError("Failed to browse target DWG", error);
    toast("Unable to select DWG path.");
  }
}

async function clearLightingScheduleTargetDwg() {
  const project = getActiveLightingScheduleProject();
  const schedule = getActiveLightingSchedule();
  if (!project || !schedule) return;
  updateLightingScheduleTargetDwgPath(project, schedule, "");
  lightingScheduleSyncStatusMessage = "";
  renderLightingScheduleSyncControls(project, schedule);
  await persistActiveLightingSchedule({
    reason: "target DWG cleared",
    showStatus: true,
  });
}

async function pullLightingScheduleFromAutoCAD() {
  const project = getActiveLightingScheduleProject();
  const schedule = getActiveLightingSchedule();
  if (!project || !schedule) {
    toast("Select a project first.");
    return;
  }
  if (lightingScheduleSyncDirty) {
    const overwrite = confirm(
      "You have unsaved desktop edits.\n\nPress OK to discard them and reload the central schedule."
    );
    if (!overwrite) return;
  }
  await loadLightingScheduleFromCentralStore(project, { quiet: false });
  toast("Lighting schedule reloaded from the central store.");
}

async function pushLightingScheduleToAutoCAD() {
  await persistActiveLightingSchedule({
    reason: "manual save",
    showStatus: true,
    showToast: true,
  });
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
  lightingScheduleSyncDirty = true;
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
        lightingScheduleSyncDirty = true;
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

async function setLightingScheduleProject(index) {
  if (!Number.isInteger(index) || !db[index]) {
    lightingScheduleProjectIndex = null;
    lightingScheduleSyncStatusMessage = "";
    lightingScheduleSyncDirty = false;
    renderLightingSchedule(null);
    return;
  }
  lightingScheduleProjectIndex = index;
  lightingScheduleSyncStatusMessage = "";
  lightingScheduleSyncDirty = false;
  const select = document.getElementById("lightingScheduleProjectSelect");
  if (select) select.value = String(index);
  const { schedule, created } = ensureLightingSchedule(db[index]);
  if (created) save();
  renderLightingSchedule(schedule);
  await loadLightingScheduleFromCentralStore(db[index], { quiet: false });
}

function addLightingScheduleRow() {
  const schedule = getActiveLightingSchedule();
  if (!schedule) return;
  schedule.rows.push(createLightingScheduleRow());
  renderLightingScheduleRows(schedule);
  lightingScheduleSyncDirty = true;
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
  lightingScheduleSyncDirty = true;
  debouncedSaveLightingSchedule();
}

async function openLightingSchedule() {
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
    await loadLightingScheduleFromCentralStore(activeProject, { quiet: true });
  }

  const templateSearch = document.getElementById("lightingTemplateSearch");
  if (templateSearch) templateSearch.value = lightingTemplateQuery;
  renderLightingTemplateList(lightingTemplateQuery);
  if (!dlg.open) dlg.showModal();
  startLightingScheduleSyncWatcher();
}

function closeLightingSchedule() {
  const dlg = document.getElementById("lightingScheduleDlg");
  persistActiveLightingSchedule({ reason: "dialog close" });
  stopLightingScheduleSyncWatcher();
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
  if (String(userSettings.userName || "").trim()) return false;
  return !hasExistingUserData();
}

function hasExistingUserData() {
  if (Array.isArray(db) && db.length > 0) return true;
  if (notesDb && typeof notesDb === "object" && Object.keys(notesDb).length > 0) {
    return true;
  }
  if (
    timesheetDb &&
    typeof timesheetDb === "object" &&
    timesheetDb.weeks &&
    Object.keys(timesheetDb.weeks).length > 0
  ) {
    return true;
  }
  if (
    checklistsDb &&
    typeof checklistsDb === "object" &&
    Array.isArray(checklistsDb.checklists) &&
    checklistsDb.checklists.length > 0
  ) {
    return true;
  }
  if (googleAuthState.signedIn) return true;
  if (String(userSettings.apiKey || "").trim()) return true;
  if (String(userSettings.autocadPath || "").trim()) return true;
  return false;
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
    const saved = await persistUserSettingsLocally();
    if (!saved) {
      throw new Error("Failed to save settings.");
    }
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
      await persistUserSettingsLocally({ silent: true });
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
  const checklistResults = document.getElementById("checklistSearchResults");

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

    if (tab !== "texts" && notesResults) notesResults.innerHTML = "";
    if (tab !== "texts" && checklistResults) checklistResults.innerHTML = "";

    if (tab === "texts") {
      renderTextsView();
    } else if (tab === "tools") {
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

function initToolCardDetailsToggles() {
  document.querySelectorAll("#tools-panel .tool-card").forEach((card, index) => {
    if (card.dataset.detailsToggleReady === "true") return;

    const body = card.querySelector(".tool-card-body");
    const statusEl = card.querySelector(".tool-card-status");
    if (!body || !statusEl) return;

    const detailsId =
      body.id || `${card.id || `tool-card-${index + 1}`}-details`;
    body.id = detailsId;

    const footer = document.createElement("div");
    footer.className = "tool-card-footer";

    const toggle = document.createElement("button");
    toggle.type = "button";
    toggle.className = "tool-card-toggle";
    toggle.textContent = "Details";
    toggle.setAttribute("aria-controls", detailsId);
    toggle.setAttribute("aria-expanded", "false");

    const setExpanded = (expanded) => {
      card.classList.toggle("details-expanded", expanded);
      toggle.textContent = expanded ? "Hide details" : "Details";
      toggle.setAttribute("aria-expanded", String(expanded));
    };

    setExpanded(false);

    toggle.addEventListener("click", (e) => {
      e.stopPropagation();
      setExpanded(!card.classList.contains("details-expanded"));
    });

    toggle.addEventListener("keydown", (e) => {
      e.stopPropagation();
    });

    footer.append(toggle, statusEl);
    card.appendChild(footer);
    card.dataset.detailsToggleReady = "true";
  });
}

function guardUnderConstructionToolAccess(card, label, event) {
  if (
    !card ||
    card.dataset.underConstruction !== "true" ||
    areUnderConstructionToolsEnabled()
  ) {
    return true;
  }

  if (event?.type === "keydown" && event.key !== "Enter" && event.key !== " ") {
    return false;
  }

  event?.preventDefault();
  event?.stopPropagation();
  toast(
    `${label} is under construction. Enable preview access in Settings to use it.`,
    4000
  );
  return false;
}

function initEventListeners() {
  document.getElementById("search").addEventListener(
    "input",
    debounce(() => render(), 250)
  );
  document.getElementById("notesSearch").addEventListener(
    "input",
    debounce((e) => handleNotebookSearchInput(e), 250)
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
  const deliverableNotepadAddBtn = document.getElementById(
    "deliverableNotepadAddBtn"
  );
  if (deliverableNotepadAddBtn) {
    deliverableNotepadAddBtn.onclick = () => addDeliverablesToNotepadSelection();
  }
  const deliverableNotepadRemoveBtn = document.getElementById(
    "deliverableNotepadRemoveBtn"
  );
  if (deliverableNotepadRemoveBtn) {
    deliverableNotepadRemoveBtn.onclick = () =>
      removeDeliverablesFromNotepadSelection();
  }
  const deliverableNotepadExportBtn = document.getElementById(
    "deliverableNotepadExportBtn"
  );
  if (deliverableNotepadExportBtn) {
    deliverableNotepadExportBtn.onclick = () =>
      exportSelectedDeliverablesToExcel();
  }
  const copyProjectLocallyFolderList = document.getElementById(
    "copyProjectLocallyFolderList"
  );
  if (copyProjectLocallyFolderList) {
    copyProjectLocallyFolderList.addEventListener("change", (event) => {
      if (
        event.target?.matches?.(
          'input[type="checkbox"][data-copy-project-parent-checkbox][data-folder-id]'
        )
      ) {
        handleCopyProjectLocallyParentCheckboxChange(
          String(event.target.dataset.folderId || "").trim(),
          event.target.checked === true
        );
        return;
      }
      if (
        event.target?.matches?.(
          'input[type="checkbox"][data-copy-project-child-checkbox][data-folder-id][data-child-id]'
        )
      ) {
        handleCopyProjectLocallyChildCheckboxChange(
          String(event.target.dataset.folderId || "").trim(),
          String(event.target.dataset.childId || "").trim(),
          event.target.checked === true
        );
      }
    });
    copyProjectLocallyFolderList.addEventListener("click", async (event) => {
      const expandBtn = event.target?.closest?.(
        'button[data-copy-project-expand-btn][data-folder-id]'
      );
      if (!expandBtn) return;
      event.preventDefault();
      await toggleCopyProjectLocallyFolderChildren(
        String(expandBtn.dataset.folderId || "").trim()
      );
    });
    copyProjectLocallyFolderList.addEventListener("keydown", async (event) => {
      const expandBtn = event.target?.closest?.(
        'button[data-copy-project-expand-btn][data-folder-id]'
      );
      if (!expandBtn) return;
      if (event.key !== "Enter" && event.key !== " ") return;
      event.preventDefault();
      await toggleCopyProjectLocallyFolderChildren(
        String(expandBtn.dataset.folderId || "").trim()
      );
    });
    copyProjectLocallyFolderList.addEventListener("keydown", (event) => {
      if (
        event.target?.matches?.(
          'input[type="checkbox"][data-copy-project-parent-checkbox][data-folder-id]'
        ) ||
        event.target?.matches?.(
          'input[type="checkbox"][data-copy-project-child-checkbox][data-folder-id][data-child-id]'
        )
      ) {
        if (event.key === "Enter") {
          event.preventDefault();
          event.target.click();
        }
      }
    });
  }
  const localProjectManagerCopyTabBtn = document.getElementById(
    "localProjectManagerCopyTabBtn"
  );
  if (localProjectManagerCopyTabBtn) {
    localProjectManagerCopyTabBtn.onclick = () => switchLocalProjectManagerTab("copy");
  }
  const localProjectManagerSyncTabBtn = document.getElementById(
    "localProjectManagerSyncTabBtn"
  );
  if (localProjectManagerSyncTabBtn) {
    localProjectManagerSyncTabBtn.onclick = () => switchLocalProjectManagerTab("sync");
  }
  const copyProjectLocallyChooseFolderBtn = document.getElementById(
    "copyProjectLocallyChooseFolderBtn"
  );
  if (copyProjectLocallyChooseFolderBtn) {
    copyProjectLocallyChooseFolderBtn.onclick = async () => {
      try {
        await chooseAndPreviewCopyProjectLocallyServerFolder();
      } catch (error) {
        toast(error?.message || "Failed to load project folders.", 7000);
      }
    };
  }
  const copyProjectLocallyOpenExistingBtn = document.getElementById(
    "copyProjectLocallyOpenExistingBtn"
  );
  if (copyProjectLocallyOpenExistingBtn) {
    copyProjectLocallyOpenExistingBtn.onclick = async () => {
      const localProjectPath = copyProjectLocallyDialogState.localProjectPath || "";
      if (!localProjectPath) return;
      await window.pywebview.api.open_path(localProjectPath);
    };
  }
  const copyProjectLocallySelectDefaultsBtn = document.getElementById(
    "copyProjectLocallySelectDefaultsBtn"
  );
  if (copyProjectLocallySelectDefaultsBtn) {
    copyProjectLocallySelectDefaultsBtn.onclick = () =>
      applyCopyProjectLocallySelectionMode("defaults");
  }
  const copyProjectLocallySelectAllBtn = document.getElementById(
    "copyProjectLocallySelectAllBtn"
  );
  if (copyProjectLocallySelectAllBtn) {
    copyProjectLocallySelectAllBtn.onclick = () =>
      applyCopyProjectLocallySelectionMode("all");
  }
  const copyProjectLocallyClearAllBtn = document.getElementById(
    "copyProjectLocallyClearAllBtn"
  );
  if (copyProjectLocallyClearAllBtn) {
    copyProjectLocallyClearAllBtn.onclick = () =>
      applyCopyProjectLocallySelectionMode("none");
  }
  const copyProjectLocallyConfirmBtn = document.getElementById(
    "copyProjectLocallyConfirmBtn"
  );
  if (copyProjectLocallyConfirmBtn) {
    copyProjectLocallyConfirmBtn.onclick = () => {
      const selectionPayload = buildCopyProjectLocallySelectionPayload();
      if (!selectionPayload.hasSelection) {
        toast("Select at least one folder to copy locally.");
        updateCopyProjectLocallyDialogSummary();
        return;
      }
      const dialog = document.getElementById("copyProjectLocallyDlg");
      if (dialog) {
        dialog.dataset.localProjectManagerAction = "copy";
        dialog.close("confirm");
      }
    };
  }
  const localProjectManagerSyncSearch = document.getElementById(
    "localProjectManagerSyncSearch"
  );
  if (localProjectManagerSyncSearch) {
    localProjectManagerSyncSearch.addEventListener("input", (event) => {
      copyProjectLocallyDialogState.sync.searchQuery = String(
        event.target?.value || ""
      ).trim();
      renderCopyProjectLocallyDialog();
    });
  }
  const localProjectManagerManualLocalPath = document.getElementById(
    "localProjectManagerManualLocalPath"
  );
  if (localProjectManagerManualLocalPath) {
    localProjectManagerManualLocalPath.addEventListener("input", (event) => {
      copyProjectLocallyDialogState.sync.manualLocalProjectPath = normalizeWindowsPath(
        event.target?.value || ""
      );
    });
    localProjectManagerManualLocalPath.addEventListener("keydown", (event) => {
      if (event.key !== "Enter") return;
      event.preventDefault();
      applyLocalProjectManagerManualLocalPath();
    });
  }
  const localProjectManagerBrowseLocalPathBtn = document.getElementById(
    "localProjectManagerBrowseLocalPathBtn"
  );
  if (localProjectManagerBrowseLocalPathBtn) {
    localProjectManagerBrowseLocalPathBtn.onclick = async () => {
      try {
        await browseLocalProjectManagerLocalPath();
      } catch (error) {
        toast(error?.message || "Failed to choose a local project folder.", 7000);
      }
    };
  }
  const localProjectManagerUseManualLocalPathBtn = document.getElementById(
    "localProjectManagerUseManualLocalPathBtn"
  );
  if (localProjectManagerUseManualLocalPathBtn) {
    localProjectManagerUseManualLocalPathBtn.onclick = () =>
      applyLocalProjectManagerManualLocalPath();
  }
  const localProjectManagerManualServerPath = document.getElementById(
    "localProjectManagerManualServerPath"
  );
  if (localProjectManagerManualServerPath) {
    localProjectManagerManualServerPath.addEventListener("input", (event) => {
      copyProjectLocallyDialogState.sync.manualServerProjectPath = normalizeWindowsPath(
        event.target?.value || ""
      );
      resetLocalProjectManagerSyncPreview();
      renderCopyProjectLocallyDialog();
    });
  }
  const localProjectManagerBrowseServerPathBtn = document.getElementById(
    "localProjectManagerBrowseServerPathBtn"
  );
  if (localProjectManagerBrowseServerPathBtn) {
    localProjectManagerBrowseServerPathBtn.onclick = async () => {
      try {
        await browseLocalProjectManagerServerPath();
      } catch (error) {
        toast(error?.message || "Failed to choose a server project folder.", 7000);
      }
    };
  }
  const localProjectManagerPreviewSyncBtn = document.getElementById(
    "localProjectManagerPreviewSyncBtn"
  );
  if (localProjectManagerPreviewSyncBtn) {
    localProjectManagerPreviewSyncBtn.onclick = async () => {
      await previewLocalProjectManagerSync();
    };
  }
  const localProjectManagerSyncProjectList = document.getElementById(
    "localProjectManagerSyncProjectList"
  );
  if (localProjectManagerSyncProjectList) {
    localProjectManagerSyncProjectList.addEventListener("click", (event) => {
      const row = event.target?.closest?.(
        "button[data-local-project-path]"
      );
      if (!row) return;
      const targetPath = normalizeWindowsPath(row.dataset.localProjectPath || "");
      const entry = copyProjectLocallyDialogState.sync.projectEntries.find(
        (item) =>
          normalizeWindowsPath(item?.localProjectPath || "").toLowerCase() ===
          targetPath.toLowerCase()
      );
      if (entry) {
        selectLocalProjectManagerProjectEntry(entry);
      }
    });
  }
  const localProjectManagerSyncRecommendationList = document.getElementById(
    "localProjectManagerSyncRecommendationList"
  );
  if (localProjectManagerSyncRecommendationList) {
    localProjectManagerSyncRecommendationList.addEventListener("change", (event) => {
      if (
        !event.target?.matches?.(
          'input[type="checkbox"][data-local-project-manager-sync-checkbox][data-relative-path]'
        )
      ) {
        return;
      }
      const relativePath = String(event.target.dataset.relativePath || "").trim();
      const candidate = copyProjectLocallyDialogState.sync.candidateFiles.find(
        (entry) => entry.relativePath === relativePath
      );
      if (!candidate) return;
      candidate.selected = event.target.checked === true;
      renderCopyProjectLocallyDialog();
    });
  }
  const localProjectManagerSyncSelectAllBtn = document.getElementById(
    "localProjectManagerSyncSelectAllBtn"
  );
  if (localProjectManagerSyncSelectAllBtn) {
    localProjectManagerSyncSelectAllBtn.onclick = () => {
      copyProjectLocallyDialogState.sync.candidateFiles.forEach((entry) => {
        entry.selected = true;
      });
      renderCopyProjectLocallyDialog();
    };
  }
  const localProjectManagerSyncClearAllBtn = document.getElementById(
    "localProjectManagerSyncClearAllBtn"
  );
  if (localProjectManagerSyncClearAllBtn) {
    localProjectManagerSyncClearAllBtn.onclick = () => {
      copyProjectLocallyDialogState.sync.candidateFiles.forEach((entry) => {
        entry.selected = false;
      });
      renderCopyProjectLocallyDialog();
    };
  }
  const localProjectManagerSyncConfirmBtn = document.getElementById(
    "localProjectManagerSyncConfirmBtn"
  );
  if (localProjectManagerSyncConfirmBtn) {
    localProjectManagerSyncConfirmBtn.onclick = () => {
      const payload = buildLocalProjectManagerSyncConfirmPayload();
      if (!payload?.selectedRelativePaths?.length) {
        toast("Select at least one file to sync to the server.");
        updateLocalProjectManagerSyncSummary();
        return;
      }
      const dialog = document.getElementById("copyProjectLocallyDlg");
      if (dialog) {
        dialog.dataset.localProjectManagerAction = "sync";
        dialog.close("confirm");
      }
    };
  }

  document.querySelectorAll(".tool-card-settings").forEach((btn) => {
    btn.replaceChildren(createSettingsIcon(14));
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

  initToolCardDetailsToggles();

  const handleAppFocus = async () => {
    if (!googleAuthBusy) {
      await loadGoogleAuthState({ silent: true });
    }
    checkBundlesForUpdates({ showIndicator: true });
  };
  window.addEventListener("focus", handleAppFocus);
  document.addEventListener("visibilitychange", () => {
    if (!document.hidden) handleAppFocus();
  });

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
  initExpenseImagePreviewDialog();

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
  const headerGoogleAuthBtn = document.getElementById("headerGoogleAuthBtn");
  if (headerGoogleAuthBtn) {
    headerGoogleAuthBtn.onclick = () => handleHeaderGoogleAuthAction();
  }
  document.getElementById("settingsBtn").onclick = async () => {
    hideSetupHelpBanner(); // Hide the banner when user manually opens settings
    await populateSettingsModal();
    document.getElementById("settingsDlg").showModal();
  };
  document.getElementById("statsBtn").onclick = () => showStatsModal();
  document.getElementById("settings_howToSetupBtn").onclick = () =>
    document.getElementById("apiKeyHelpDlg").showModal();
  const settingsGoogleAuthActionBtn = document.getElementById(
    "settings_googleAuthActionBtn"
  );
  if (settingsGoogleAuthActionBtn) {
    settingsGoogleAuthActionBtn.onclick = () => handleGoogleAuthAction();
  }
  const googleSignOutBtn = document.getElementById("googleSignOutBtn");
  if (googleSignOutBtn) {
    googleSignOutBtn.onclick = () => handleGoogleSignOut();
  }
  const outlookScanBtn = document.getElementById("outlookScanBtn");
  if (outlookScanBtn) {
    outlookScanBtn.onclick = () => openOutlookScanDialog();
  }
  const emailIntakeModePaste = document.getElementById("emailIntakeModePaste");
  if (emailIntakeModePaste) {
    emailIntakeModePaste.onchange = (event) => {
      if (event?.target?.checked) setEmailIntakeMode("paste");
    };
  }
  const emailIntakeModeScan = document.getElementById("emailIntakeModeScan");
  if (emailIntakeModeScan) {
    emailIntakeModeScan.onchange = (event) => {
      if (event?.target?.checked) setEmailIntakeMode("scan");
    };
  }
  const btnProcessEmail = document.getElementById("btnProcessEmail");
  if (btnProcessEmail) {
    btnProcessEmail.onclick = () => processEmailIntakePaste();
  }
  const outlookScanRunBtn = document.getElementById("outlookScanRunBtn");
  if (outlookScanRunBtn) {
    outlookScanRunBtn.onclick = () => runOutlookInboxScan();
  }
  const outlookScanDate = document.getElementById("outlookScanDate");
  if (outlookScanDate) {
    outlookScanDate.onchange = (event) =>
      setOutlookScanDate(event?.target?.value || getTodayLocalDateInputValue());
  }
  const outlookScanReport = document.getElementById("outlookScanReport");
  if (outlookScanReport) {
    outlookScanReport.addEventListener("toggle", () => {
      outlookScanState.reportOpen = !!outlookScanReport.open;
    });
  }
  const headerAccountSignOutBtn = document.getElementById(
    "headerAccountSignOutBtn"
  );
  if (headerAccountSignOutBtn) {
    headerAccountSignOutBtn.onclick = () => handleGoogleSignOut();
  }
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
      if (openAttachmentPanelContext?.trigger?.closest("#editDlg")) {
        void requestAttachmentPanelClose();
      }
      if (modalEmailSession.active) flushModalEmailSession(false);
    });
  }

  const pathInput = document.getElementById("f_path");
  if (pathInput) {
    pathInput.addEventListener("input", () => {
      pathInput.title = pathInput.value;
      applyProjectFromPath(pathInput.value);
    });
    pathInput.addEventListener("paste", () => {
      setTimeout(
        () => normalizeProjectPathInput(pathInput, { forceProjectFields: true }),
        0
      );
    });
    pathInput.addEventListener("blur", () => {
      normalizeProjectPathInput(pathInput);
    });
  }

  document.querySelectorAll(".projects-filter-dropdown").forEach((dropdown) => {
    const filterKey = dropdown.dataset.filterDropdown;
    const trigger = dropdown.querySelector(".projects-filter-trigger");
    const menu = dropdown.querySelector(".projects-filter-menu");

    trigger?.addEventListener("click", (e) => {
      e.stopPropagation();
      const willOpen = !dropdown.classList.contains("open");
      setProjectsFilterDropdownState(dropdown, willOpen, {
        focusSelected: willOpen
      });
    });

    trigger?.addEventListener("keydown", (e) => {
      if (e.key !== "ArrowDown" && e.key !== "ArrowUp") return;
      e.preventDefault();
      setProjectsFilterDropdownState(dropdown, true, { focusSelected: true });
    });

    menu?.addEventListener("click", (e) => {
      const option = e.target.closest(".projects-filter-option");
      if (!option) return;

      e.stopPropagation();
      setProjectsFilterValue(filterKey, option.dataset.filterValue || "all");
      setProjectsFilterDropdownState(dropdown, false);
      render();
      trigger?.focus();
    });

    menu?.addEventListener("keydown", (e) => {
      const option = e.target.closest(".projects-filter-option");
      if (!option) return;

      if (e.key === "ArrowDown") {
        e.preventDefault();
        moveProjectsFilterOptionFocus(dropdown, option, 1);
        return;
      }

      if (e.key === "ArrowUp") {
        e.preventDefault();
        moveProjectsFilterOptionFocus(dropdown, option, -1);
        return;
      }

      if (e.key === "Home") {
        e.preventDefault();
        getProjectsFilterOptions(dropdown)[0]?.focus();
        return;
      }

      if (e.key === "End") {
        e.preventDefault();
        const options = getProjectsFilterOptions(dropdown);
        options[options.length - 1]?.focus();
        return;
      }

      if (e.key === "Escape") {
        e.preventDefault();
        setProjectsFilterDropdownState(dropdown, false);
        trigger?.focus();
      }
    });
  });

  document.addEventListener("click", (e) => {
    if (!openProjectsFilterDropdown) return;
    if (openProjectsFilterDropdown.contains(e.target)) return;
    setProjectsFilterDropdownState(openProjectsFilterDropdown, false);
  });

  document.addEventListener("click", (e) => {
    if (!headerAccountPopoverOpen) return;
    const dropdown = document.getElementById("headerAccountDropdown");
    if (dropdown?.contains(e.target)) return;
    setHeaderAccountPopoverOpen(false);
  });

  document.addEventListener("keydown", (e) => {
    if (e.key !== "Escape" || !openProjectsFilterDropdown) return;
    const dropdown = openProjectsFilterDropdown;
    setProjectsFilterDropdownState(dropdown, false);
    dropdown.querySelector(".projects-filter-trigger")?.focus();
  });

  document.addEventListener("keydown", (e) => {
    if (e.key !== "Escape" || !headerAccountPopoverOpen) return;
    e.preventDefault();
    setHeaderAccountPopoverOpen(false, { focusTrigger: true });
  });

  syncProjectsFilterDropdowns();

  document
    .getElementById("toolPublishDwgs")
    .addEventListener("click", async (e) => {
      const launchContext = resolveCadLaunchContextForTool();
      console.debug("Workroom CAD launch context (publish):", launchContext);
      if (e.currentTarget.classList.contains("running")) return;
      if (!userSettings.autocadPath) {
        await showAutocadSelectModal();
        return;
      }
      const activityId = beginActivity({
        toolId: "toolPublishDwgs",
        message: "Initializing...",
        progress: 5,
      });
      try {
        const result = launchContext
          ? await window.pywebview.api.run_publish_script(launchContext, activityId)
          : await window.pywebview.api.run_publish_script(null, activityId);
        if (result?.status === "error") {
          failActivity(activityId, {
            message: result.message || "Failed to start Publish DWGs.",
          });
        }
      } catch (error) {
        failActivity(activityId, {
          message: error?.message || "Failed to start Publish DWGs.",
        });
      }
    });

  document
    .getElementById("toolFreezeLayers")
    .addEventListener("click", async (e) => {
      const launchContext = resolveCadLaunchContextForTool();
      console.debug("Workroom CAD launch context (freeze):", launchContext);
      if (e.currentTarget.classList.contains("running")) return;
      if (!userSettings.autocadPath) {
        await showAutocadSelectModal();
        return;
      }
      const activityId = beginActivity({
        toolId: "toolFreezeLayers",
        message: "Initializing...",
        progress: 5,
      });
      try {
        const result = launchContext
          ? await window.pywebview.api.run_freeze_layers_script(
              launchContext,
              activityId
            )
          : await window.pywebview.api.run_freeze_layers_script(null, activityId);
        if (result?.status === "error") {
          failActivity(activityId, {
            message: result.message || "Failed to start Freeze Layers.",
          });
        }
      } catch (error) {
        failActivity(activityId, {
          message: error?.message || "Failed to start Freeze Layers.",
        });
      }
    });

  document
    .getElementById("toolThawLayers")
    .addEventListener("click", async (e) => {
      const launchContext = resolveCadLaunchContextForTool();
      console.debug("Workroom CAD launch context (thaw):", launchContext);
      if (e.currentTarget.classList.contains("running")) return;
      if (!userSettings.autocadPath) {
        await showAutocadSelectModal();
        return;
      }
      const activityId = beginActivity({
        toolId: "toolThawLayers",
        message: "Initializing...",
        progress: 5,
      });
      try {
        const result = launchContext
          ? await window.pywebview.api.run_thaw_layers_script(launchContext, activityId)
          : await window.pywebview.api.run_thaw_layers_script(null, activityId);
        if (result?.status === "error") {
          failActivity(activityId, {
            message: result.message || "Failed to start Thaw Layers.",
          });
        }
      } catch (error) {
        failActivity(activityId, {
          message: error?.message || "Failed to start Thaw Layers.",
        });
      }
    });

  document
    .getElementById("toolCleanXrefs")
    .addEventListener("click", async (e) => {
      const launchContext = resolveCadLaunchContextForTool();
      if (e.currentTarget.classList.contains("running")) return;
      if (!userSettings.autocadPath) {
        await showAutocadSelectModal();
        return;
      }
      const activityId = beginActivity({
        toolId: "toolCleanXrefs",
        message: "Initializing...",
        progress: 5,
      });
      try {
        const result = launchContext
          ? await window.pywebview.api.run_clean_xrefs_script(launchContext, activityId)
          : await window.pywebview.api.run_clean_xrefs_script(null, activityId);
        if (result?.status === "error") {
          failActivity(activityId, {
            message: result.message || "Failed to start Clean Xrefs.",
          });
        }
      } catch (error) {
        failActivity(activityId, {
          message: error?.message || "Failed to start Clean Xrefs.",
        });
      }
    });

  const bindTemplateToolButton = (toolId, templateKey, label) => {
    const button = document.getElementById(toolId);
    if (!button) return;
    const handler = async () => {
      const launchContext = resolveCadLaunchContextForTool();
      console.debug(`Workroom launch context (${templateKey} template):`, launchContext);
      await handleTemplateToolSave(templateKey, label, {
        launchContext,
        toolId,
      });
    };
    button.addEventListener("click", handler);
    button.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        handler();
      }
    });
  };

  bindTemplateToolButton(
    "toolCreateNarrativeTemplate",
    "narrative",
    "Narrative of Changes"
  );
  bindTemplateToolButton(
    "toolCreatePlanCheckTemplate",
    "planCheck",
    "Plan Check Comments"
  );
  const backupDrawingsBtn = document.getElementById("toolBackupDrawings");
  const generalToolsGrid = Array.from(document.querySelectorAll(".tools-section"))
    .find(
      (section) => section.querySelector(".tools-section-title")?.textContent?.trim() === "General"
    )
    ?.querySelector(".tools-grid");
  if (generalToolsGrid) {
    const generalToolOrder = [
      "toolPublishDwgs",
      "toolFreezeLayers",
      "toolThawLayers",
      "toolCleanXrefs",
      "toolCopyProjectLocally",
      "toolBackupDrawings",
    ];
    generalToolOrder
      .map((toolId) => document.getElementById(toolId))
      .filter(Boolean)
      .forEach((toolCard) => {
        generalToolsGrid.appendChild(toolCard);
      });
  }

  if (backupDrawingsBtn) {
    const handler = async () => {
      if (backupDrawingsBtn.classList.contains("running")) return;
      const activityId = beginActivity({
        toolId: "toolBackupDrawings",
        message: "Resolving project folder...",
        progress: 8,
      });
      const launchContext = resolveCadLaunchContextForTool();
      const launchSource = String(launchContext?.source || "").trim().toLowerCase();
      const hasContextProjectPath = hasLaunchContextProjectPath(launchContext);

      try {
        const selectProjectFolder = async () => {
          updateActivity(activityId, {
            message: "Select project folder...",
            progress: 10,
          });
          const selection = await window.pywebview.api.select_folder(
            getLaunchContextProjectRoot(launchContext) || null
          );
          if (selection?.status === "error") {
            throw new Error(selection.message || "Failed to choose a folder.");
          }
          if (!selection || selection.status === "cancelled" || !selection.path) {
            return "";
          }
          return String(selection.path || "").trim();
        };

        let result = null;
        let selectedProjectPath = "";

        if (launchSource === "workroom" || hasContextProjectPath) {
          updateActivity(activityId, {
            message: "Resolving project folder...",
            progress: 18,
          });
          result = await window.pywebview.api.backup_project_drawings(null, launchContext);

          const resultCode = String(result?.code || "").trim().toLowerCase();
          const resultMessage = String(result?.message || "").trim().toLowerCase();
          const needsManualSelection =
            result?.status !== "success" &&
            (resultCode === "manual_selection_required" ||
              resultCode === "server_project_path_required" ||
              resultMessage.includes("path does not exist"));

          if (needsManualSelection) {
            updateActivity(activityId, {
              message: "Could not auto-resolve project folder. Select it manually...",
              progress: 10,
            });
            selectedProjectPath = await selectProjectFolder();
            if (!selectedProjectPath) {
              result = null;
            }
          }
        } else {
          selectedProjectPath = await selectProjectFolder();
        }

        if (selectedProjectPath) {
          updateActivity(activityId, {
            message: "Creating archive backup...",
            progress: 42,
          });
          result = await window.pywebview.api.backup_project_drawings(
            selectedProjectPath,
            launchContext
          );
        }

        if (!result && !selectedProjectPath) {
          acceptActivity(activityId);
          return;
        }

        if (result?.status !== "success") {
          throw new Error(result?.message || "Failed to create drawing backup.");
        }

        const failedFiles = Array.isArray(result?.failedFiles) ? result.failedFiles : [];
        const failedFileCount =
          Number.isFinite(Number(result?.failedFileCount))
            ? Number(result.failedFileCount)
            : failedFiles.length;
        const missingFolders = Array.isArray(result?.missingSourceFolders)
          ? result.missingSourceFolders
          : [];
        const warningParts = [];
        if (missingFolders.length) {
          warningParts.push(
            `${missingFolders.length} missing folder${missingFolders.length === 1 ? "" : "s"}`
          );
        }
        if (failedFileCount > 0) {
          warningParts.push(
            `${failedFileCount} failed file${failedFileCount === 1 ? "" : "s"}`
          );
        }
        const hasWarnings = warningParts.length > 0;

        completeActivity(activityId, {
          status: hasWarnings ? ACTIVITY_STATUS.WARNING : ACTIVITY_STATUS.SUCCESS,
          message: hasWarnings
            ? `Drawing backup created with warnings: ${warningParts.join(", ")}.`
            : "Drawing backup created.",
          openFolderPath: String(result?.archivePath || "").trim(),
        });
      } catch (e) {
        const message = e?.message || "Failed to create drawing backup.";
        failActivity(activityId, { message });
      } finally {
        backupDrawingsBtn.classList.remove("running");
      }
    };

    backupDrawingsBtn.addEventListener("click", handler);
    backupDrawingsBtn.addEventListener("keydown", (e) => {
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
      const activityId = beginActivity({
        toolId: "toolCopyProjectLocally",
        message: "Preparing Local Project Manager...",
        progress: 8,
      });
      const launchContext = resolveCadLaunchContextForTool();
      const launchSource = String(launchContext?.source || "").trim().toLowerCase();
      const hasContextProjectPath = hasLaunchContextProjectPath(launchContext);

      try {
        const syncWorkroomLocalProject = async (localProjectPath) => {
          if (launchSource !== "workroom" || !localProjectPath) return;
          const activeProject = db[activeChecklistProject];
          if (!activeProject) return;
          const changed = await setWorkroomLocalProjectPath(
            activeProject,
            localProjectPath,
            { saveNow: true }
          );
          if (changed) {
            renderWorkroomProjectHeader();
          }
        };

        let preview = null;
        if (launchSource === "workroom" || hasContextProjectPath) {
          updateActivity(activityId, {
            message: "Resolving project folder...",
            progress: 18,
          });
          preview = await window.pywebview.api.preview_copy_project_locally(
            null,
            launchContext
          );
        }

        updateActivity(activityId, {
          message: "Waiting for selection...",
          progress: 20,
        });
        const managerResult = await openCopyProjectLocallyDialog(
          preview,
          launchContext
        );
        if (!managerResult) {
          acceptActivity(activityId);
          return;
        }

        if (managerResult.action === "sync") {
          updateActivity(activityId, {
            message: "Syncing selected files...",
            progress: 42,
          });
          const syncResult = await window.pywebview.api.apply_local_project_manager_sync(
            managerResult.localProjectPath,
            managerResult.serverProjectPath,
            managerResult.selectedRelativePaths,
            launchContext
          );

          if (syncResult?.status !== "success") {
            throw new Error(syncResult?.message || "Failed to sync local files to the server.");
          }

          const failedFiles = Array.isArray(syncResult?.failedFiles)
            ? syncResult.failedFiles
            : [];
          const blockedEntries = Array.isArray(syncResult?.blockedEntries)
            ? syncResult.blockedEntries
            : [];
          const copiedFileCount = Number.isFinite(Number(syncResult?.copiedFileCount))
            ? Number(syncResult.copiedFileCount)
            : 0;
          const failedFileCount = Number.isFinite(Number(syncResult?.failedFileCount))
            ? Number(syncResult.failedFileCount)
            : failedFiles.length;
          const hasWarnings = failedFileCount > 0 || blockedEntries.length > 0;

          const warningParts = [];
          if (failedFileCount > 0) {
            warningParts.push(
              `${failedFileCount} failed file${failedFileCount === 1 ? "" : "s"}`
            );
          }
          if (blockedEntries.length > 0) {
            warningParts.push(
              `${blockedEntries.length} blocked entr${blockedEntries.length === 1 ? "y" : "ies"}`
            );
          }

          completeActivity(activityId, {
            status: hasWarnings ? ACTIVITY_STATUS.WARNING : ACTIVITY_STATUS.SUCCESS,
            message: hasWarnings
              ? `Server sync completed with warnings: ${warningParts.join(", ")}.`
              : `Server sync completed (${copiedFileCount} file${
                  copiedFileCount === 1 ? "" : "s"
                } copied).`,
            openFolderPath: String(syncResult?.serverProjectPath || "").trim(),
          });
          return;
        }

        const selectionPayload = managerResult.selectionPayload;
        if (!selectionPayload?.hasSelection) {
          acceptActivity(activityId);
          return;
        }

        updateActivity(activityId, {
          message: "Copying selected folders...",
          progress: 42,
        });
        const copyResult = await window.pywebview.api.copy_project_locally(
          managerResult.serverProjectPath || "",
          launchContext,
          selectionPayload.selectedFolderNames,
          selectionPayload.selectedFolderRequests
        );

        const resultCode = String(copyResult?.code || "").trim().toLowerCase();
        if (copyResult?.status !== "success") {
          if (resultCode === "local_project_exists" && copyResult?.localProjectPath) {
            await syncWorkroomLocalProject(copyResult.localProjectPath);
            completeActivity(activityId, {
              message: "Local project already exists. Linked existing copy.",
              openFolderPath: String(copyResult.localProjectPath || "").trim(),
            });
            return;
          }
          throw new Error(copyResult?.message || "Failed to copy project locally.");
        }

        const failedFiles = Array.isArray(copyResult?.failedFiles) ? copyResult.failedFiles : [];
        const failedFileCount =
          Number.isFinite(Number(copyResult?.failedFileCount))
            ? Number(copyResult.failedFileCount)
            : failedFiles.length;
        const hasWarnings = failedFileCount > 0;

        await syncWorkroomLocalProject(copyResult?.localProjectPath);

        const missingFolders = Array.isArray(copyResult?.missingServerFolders)
          ? copyResult.missingServerFolders
          : [];
        const warningParts = [];
        if (failedFileCount > 0) {
          warningParts.push(
            `${failedFileCount} failed file${failedFileCount === 1 ? "" : "s"}`
          );
        }
        if (missingFolders.length) {
          warningParts.push(
            `${missingFolders.length} missing folder${missingFolders.length === 1 ? "" : "s"}`
          );
        }

        completeActivity(activityId, {
          status:
            hasWarnings || missingFolders.length
              ? ACTIVITY_STATUS.WARNING
              : ACTIVITY_STATUS.SUCCESS,
          message:
            hasWarnings || missingFolders.length
              ? `Project copied locally with warnings: ${warningParts.join(", ")}.`
              : "Project copied locally.",
          openFolderPath: String(copyResult?.localProjectPath || "").trim(),
        });
      } catch (e) {
        const message = e?.message || "Failed to run Local Project Manager.";
        failActivity(activityId, { message });
      } finally {
        copyProjectLocallyBtn.classList.remove("running");
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
    cbBreakerDrop.addEventListener("click", () => {
      void selectCircuitBreakerImage("breaker");
    });
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
        void selectCircuitBreakerImage("breaker");
      }
    });
  }

  const cbDirectoryDrop = document.getElementById("cbDirectoryDrop");
  if (cbDirectoryDrop) {
    cbDirectoryDrop.addEventListener("click", () => {
      void selectCircuitBreakerImage("directory");
    });
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
        void selectCircuitBreakerImage("directory");
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
    const openLightingScheduleHandler = (e) => {
      if (
        !guardUnderConstructionToolAccess(
          lightingScheduleBtn,
          "Light Fixture Scheduler",
          e
        )
      ) {
        return;
      }
      openLightingSchedule();
    };
    lightingScheduleBtn.addEventListener("click", openLightingScheduleHandler);
    lightingScheduleBtn.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        openLightingScheduleHandler(e);
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
      stopLightingScheduleSyncWatcher();
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
      lightingScheduleSyncDirty = true;
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
      lightingScheduleSyncDirty = true;
      debouncedSaveLightingSchedule();
    });
  }

  const title24ComplianceBtn = document.getElementById("toolTitle24Compliance");
  if (title24ComplianceBtn) {
    const openTitle24ComplianceHandler = async (e) => {
      if (
        !guardUnderConstructionToolAccess(
          title24ComplianceBtn,
          "Title 24 Compliance",
          e
        )
      ) {
        return;
      }
      await openTitle24Compliance();
    };
    title24ComplianceBtn.addEventListener("click", (e) => {
      openTitle24ComplianceHandler(e);
    });
    title24ComplianceBtn.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        openTitle24ComplianceHandler(e);
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
        const saved = await persistUserSettingsLocally({
          skipCloud: true,
          silent: true,
        });
        if (!saved) {
          throw new Error("Failed to save settings.");
        }
      } catch (e) {
        toast("⚠️ Could not save settings.");
        return;
      }
      closeDlg("autocadSelectDlg");
    });

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
  const separateCompletionGroupsSetting = document.getElementById(
    "settings_separateDeliverableCompletionGroups"
  );
  if (separateCompletionGroupsSetting) {
    separateCompletionGroupsSetting.onchange = (e) => {
      userSettings.separateDeliverableCompletionGroups = e.target.checked;
      syncProjectViewPreferencesFromSettings();
      render();
      debouncedSaveUserSettings();
    };
  }
  const groupDeliverablesByProjectSetting = document.getElementById(
    "settings_groupDeliverablesByProject"
  );
  if (groupDeliverablesByProjectSetting) {
    groupDeliverablesByProjectSetting.onchange = (e) => {
      userSettings.groupDeliverablesByProject = e.target.checked;
      syncProjectViewPreferencesFromSettings();
      render();
      debouncedSaveUserSettings();
    };
  }
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

  ["settings_workroomAutoSelectCadFiles"]
    .map((id) => document.getElementById(id))
    .filter(Boolean)
    .forEach((checkbox) => {
      checkbox.onchange = (e) => {
        userSettings.workroomAutoSelectCadFiles = e.target.checked;
        syncWorkroomCadRoutingInputs();
        debouncedSaveUserSettings();
      };
    });

  const underConstructionToolsCheckbox = document.getElementById(
    "settings_enableUnderConstructionTools"
  );
  if (underConstructionToolsCheckbox) {
    underConstructionToolsCheckbox.onchange = (e) => {
      userSettings.enableUnderConstructionTools = e.target.checked;
      syncUnderConstructionToolsInputs();
      syncUnderConstructionToolsAvailability();
      debouncedSaveUserSettings();
    };
  }

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
    await loadGoogleAuthState({ silent: true });
    await loadUserSettings();
    initThemeFromPreferences();
    const [
      loadedDb,
      loadedNotes,
      loadedTimesheets,
      loadedTemplates,
      loadedChecklists,
    ] = await Promise.all([
      load(),
      loadNotes(),
      loadTimesheets(),
      loadTemplates(),
      loadChecklists(),
    ]);
    db = loadedDb || [];
    notesDb = loadedNotes || {};
    timesheetDb = loadedTimesheets || { weeks: {}, lastModified: null };
    templatesDb = loadedTemplates || {
      templates: [],
      defaultTemplatesInstalled: false,
      lastModified: null,
    };
    checklistsDb = loadedChecklists || {
      checklists: [],
      templateOverrides: {},
      lastModified: null,
    };
    await loadLocalSyncMetadata();
    syncAllCloudComparableFingerprints();

    if (googleAuthState.signedIn) {
      await bootstrapCloudSync({ silent: true });
    }

    if (checklistsDb.checklists.length > 0) {
      activeChecklistTabId =
        checklistsDb.checklists.find((checklist) => checklist.id === activeChecklistTabId)
          ?.id || checklistsDb.checklists[0].id;
    } else {
      activeChecklistTabId = null;
    }
    ensureActiveNotebookSelection(activeNotebookType);

    if (isNewUser()) {
      hideMainApp();
      showOnboardingModal();
    } else {
      showMainApp();
      renderNoteTabs();
      renderChecklistTabs();
      renderChecklistItems();
      renderTextsView();
      renderTemplates();
      renderTimesheets();
      render();
      updateTimesheetsDuplicateIndicator(
        getWeekEntries(formatWeekKey(currentTimesheetWeek))
      );

      if (userSettings.showSetupHelp !== false) {
        setTimeout(() => showSetupHelpBanner(), 1000);
      }
    }
    prefetchBundles();
  } finally {
    hideAppLoader();
  }
}

init();

