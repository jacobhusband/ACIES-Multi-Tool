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
const MAX_HOURS_PER_DAY = 12;

// Cache for loaded descriptions to prevent re-fetching
const DESCRIPTION_CACHE = {};
let bundlesPrefetchStarted = false;

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
  row.draggable = true;

  row.addEventListener("dragstart", (e) => {
    const handle = e.target.closest(".ts-drag-handle");
    if (!handle) {
      e.preventDefault();
      return;
    }
    timesheetDragEntryId = entryId;
    row.classList.add("ts-dragging");
    if (e.dataTransfer) {
      e.dataTransfer.effectAllowed = "move";
      e.dataTransfer.setData("text/plain", entryId);
    }
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

  row.addEventListener("dragend", () => {
    timesheetDragEntryId = null;
    clearTimesheetDragStyles();
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
  const addedIds = new Set(
    entries.map((e) => `${e.projectId}-${e.deliverableId}`)
  );

  if (!suggestions.length) {
    container.innerHTML =
      '<p class="muted">No projects due this week or next.</p>';
    return;
  }

  suggestions.forEach(({ project, deliverable, dueDate }) => {
    const isAdded = addedIds.has(`${project.id}-${deliverable.id}`);
    const card = createSuggestionCard(project, deliverable, dueDate, isAdded);
    container.appendChild(card);
  });
}

function createSuggestionCard(project, deliverable, dueDate, isAdded) {
  const card = el("div", {
    className: `ts-suggestion-card ${isAdded ? "added" : ""}`,
  });

  card.onclick = () => {
    if (isAdded) {
      toast("Already on the timesheet.");
      return;
    }
    addProjectToTimesheet(project, deliverable);
  };

  const header = el("div", { className: "ts-suggestion-header" });
  header.appendChild(
    el("span", { className: "ts-suggestion-id", textContent: project.id || "--" })
  );

  const ds = dueState(deliverable.due);
  header.appendChild(
    el("span", {
      className: `ts-suggestion-due ${ds}`,
      textContent: humanDate(deliverable.due),
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
    el("div", {
      className: "ts-suggestion-deliverable",
      textContent: deliverable.name || "Deliverable",
    })
  );

  if (isAdded) {
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

function addProjectToTimesheet(project, deliverable) {
  const weekKey = formatWeekKey(currentTimesheetWeek);
  const entries = getWeekEntries(weekKey);
  const projectId = project.id || "";
  const deliverableId = deliverable?.id || "";
  const alreadyAdded = entries.some(
    (e) => e.projectId === projectId && e.deliverableId === deliverableId
  );
  if (alreadyAdded) {
    toast("Project already on the timesheet.");
    return;
  }

  const entry = {
    id: createTimesheetEntryId(),
    projectId,
    projectName: project.nick || project.name || "",
    deliverableId,
    deliverableName: deliverable?.name || "",
    pmInitials: getDefaultPmInitials(),
    serviceDescription: deliverable?.name || "",
    function: getDisciplineFunction(),
    hours: { sun: 0, mon: 0, tue: 0, wed: 0, thu: 0, fri: 0, sat: 0 },
    mileage: 0,
    taskNumber: "",
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
    taskNumber: "",
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
          <div><strong>${project.id || "--"}</strong> - ${
          project.nick || project.name || "Unnamed"
        }</div>
          <div class="muted tiny">${deliverable.name || "Deliverable"} - Due: ${
          humanDate(deliverable.due) || "No date"
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

  if (!entries.length) {
    toast("No entries to export");
    return;
  }

  try {
    const result = await window.pywebview.api.export_timesheet_excel({
      weekKey,
      weekDisplay: formatWeekDisplay(currentTimesheetWeek),
      userName: userSettings.userName || "Employee",
      entries,
      totals: calculateWeekTotals(entries),
    });

    if (result.status === "success") {
      toast("Timesheet exported successfully");
    } else {
      throw new Error(result.message);
    }
  } catch (e) {
    console.error("Export failed:", e);
    toast("Failed to export timesheet: " + e.message);
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
let currentSort = { key: "due", dir: "desc" };
let statusFilter = "all";
let dueFilter = "all";

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
};
let hideNonPrimary = false;
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
  } catch (e) {
    console.error("Failed to load settings:", e);
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

function createCardHeader(deliverable, isPrimary) {
  const header = el("div", { className: "deliverable-card-header-new" });

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
    textContent: deliverable.name || "Deliverable"
  });
  title.appendChild(nameSpan);

  header.appendChild(title);

  // Due date badge
  if (deliverable.due) {
    const ds = dueState(deliverable.due);
    const badgeClass = `deliverable-due-badge ${ds === "overdue" ? "overdue" : ds === "dueSoon" ? "due-soon" : "ok"}`;
    const badgeText = ds === "overdue" ? "OVERDUE" : humanDate(deliverable.due).replace(/\//g, "/").toUpperCase();
    const badge = el("div", { className: badgeClass, textContent: badgeText });
    header.appendChild(badge);
  }

  return header;
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

function createStatusDropdown(deliverable, project) {
  const availableStatuses = ["Waiting", "Working", "Pending Review", "Complete", "Delivered"];
  const dropdown = el("div", { className: "deliverable-status-dropdown" });

  // Trigger button - compact plus icon
  const trigger = el("button", {
    className: "deliverable-status-trigger",
    title: "Edit statuses"
  });

  const plusSvg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  plusSvg.setAttribute("viewBox", "0 0 24 24");
  plusSvg.setAttribute("fill", "none");
  plusSvg.setAttribute("stroke", "currentColor");
  plusSvg.setAttribute("stroke-width", "2");
  plusSvg.innerHTML = '<path d="M12 5v14m-7-7h14"/>';

  trigger.appendChild(plusSvg);

  // Dropdown menu
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

    // Close other open dropdowns
    document.querySelectorAll(".deliverable-status-menu.open").forEach(m => {
      if (m !== menu) {
        m.classList.remove("open");
        m.previousElementSibling.classList.remove("open");
      }
    });
  });

  // Close dropdown when clicking outside
  document.addEventListener("click", (e) => {
    if (!dropdown.contains(e.target)) {
      menu.classList.remove("open");
      trigger.classList.remove("open");
    }
  });

  dropdown.append(trigger, menu);
  return dropdown;
}

function createNotesSection(deliverable, project) {
  const section = el("div", { className: "deliverable-notes-section" });

  // Header (collapsible)
  const header = el("div", { className: "deliverable-notes-header" });
  const title = el("div", { className: "deliverable-notes-title" });

  const chevronSvg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  chevronSvg.setAttribute("viewBox", "0 0 24 24");
  chevronSvg.setAttribute("fill", "none");
  chevronSvg.setAttribute("stroke", "currentColor");
  chevronSvg.setAttribute("stroke-width", "2");
  chevronSvg.innerHTML = '<polyline points="9 18 15 12 9 6"></polyline>';

  const titleText = el("span", { textContent: "Notes" });
  title.append(chevronSvg, titleText);
  header.appendChild(title);

  // Content (textarea)
  const content = el("div", { className: "deliverable-notes-content" });
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

  content.appendChild(textarea);

  // Toggle expand/collapse
  header.addEventListener("click", () => {
    const isExpanded = content.classList.toggle("expanded");
    header.classList.toggle("expanded", isExpanded);
  });

  section.append(header, content);
  return section;
}

function createTasksPreview(deliverable, card) {
  const container = el("div", { className: "deliverable-tasks-preview" });

  if (!deliverable.tasks || !deliverable.tasks.length) {
    return container;
  }

  const stats = getTaskCompletionStats(deliverable);
  const heading = el("div", {
    className: "deliverable-tasks-preview-heading",
    textContent: `Tasks (${stats.completed}/${stats.total}):`
  });

  const list = el("div", { className: "deliverable-tasks-preview-list" });
  const maxVisible = 3;

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
      const textSpan = el("span", { textContent: taskObj.text || "Task" });

      item.append(iconSpan, textSpan);

      // Make task clickable to toggle completion
      item.addEventListener("click", async (e) => {
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

        // Save changes
        await save();

        // Update progress bar
        const progressSection = card.querySelector(".deliverable-progress-section");
        if (progressSection) {
          const newStats = getTaskCompletionStats(deliverable);
          const barFill = progressSection.querySelector(".deliverable-progress-bar-fill");
          const progressText = progressSection.querySelector(".deliverable-progress-text");
          const percentageSpan = progressText.querySelector(".percentage");
          const detailSpan = progressText.querySelector(".detail");

          barFill.style.width = `${newStats.percentage}%`;
          percentageSpan.textContent = `${newStats.percentage}%`;
          detailSpan.textContent = newStats.total > 0 ? ` complete (${newStats.completed}/${newStats.total} tasks)` : " (no tasks)";

          // Update progress class
          progressText.className = `deliverable-progress-text ${newStats.percentage >= 50 ? "high" : newStats.percentage >= 25 ? "medium" : "low"}`;
        }

        // Update heading
        const newStats = getTaskCompletionStats(deliverable);
        heading.textContent = `Tasks (${newStats.completed}/${newStats.total}):`;
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
  };

  renderTaskList(false);

  container.append(heading, list);
  return container;
}

function setDeliverableDetailsCollapsed(card, isCollapsed, toggleBtn) {
  card.classList.toggle("details-collapsed", isCollapsed);
  const btn = toggleBtn || card.querySelector(".deliverable-details-toggle");
  if (!btn) return;
  btn.classList.toggle("expanded", !isCollapsed);
  const label = isCollapsed ? "Show details" : "Hide details";
  btn.setAttribute("aria-label", label);
  btn.setAttribute("title", label);
}

function createDetailsToggle(card) {
  const btn = el("button", {
    className: "deliverable-details-toggle",
    type: "button",
  });
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("viewBox", "0 0 24 24");
  svg.setAttribute("fill", "none");
  svg.setAttribute("stroke", "currentColor");
  svg.setAttribute("stroke-width", "2");
  svg.innerHTML = '<polyline points="6 9 12 15 18 9"></polyline>';
  btn.appendChild(svg);

  btn.onclick = (e) => {
    e.stopPropagation();
    setDeliverableDetailsCollapsed(
      card,
      !card.classList.contains("details-collapsed")
    );
  };

  setDeliverableDetailsCollapsed(
    card,
    card.classList.contains("details-collapsed"),
    btn
  );
  return btn;
}

function createCollapseButton(card) {
  const btn = el("button", { className: "deliverable-collapse-toggle" });
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("viewBox", "0 0 24 24");
  svg.setAttribute("fill", "none");
  svg.setAttribute("stroke", "currentColor");
  svg.setAttribute("stroke-width", "2");
  svg.innerHTML = '<polyline points="6 9 12 15 18 9"></polyline>';
  btn.appendChild(svg);

  btn.onclick = (e) => {
    e.stopPropagation();
    setDeliverableDetailsCollapsed(
      card,
      !card.classList.contains("details-collapsed")
    );
  };

  return btn;
}

function renderDeliverableCard(deliverable, isPrimary, project) {
  const card = el("div", {
    className: `deliverable-card-new ${isPrimary ? "is-primary" : ""} details-collapsed`
  });

  // Header: name + due badge
  const header = createCardHeader(deliverable, isPrimary);

  // Collapse toggle (only for non-primary)
  if (!isPrimary) {
    const collapseBtn = createCollapseButton(card);
    header.appendChild(collapseBtn);
  }

  // Progress: bar + percentage text
  const progress = createProgressSection(deliverable);

  // Status section: badges + dropdown inline
  const statusSection = el("div", { className: "deliverable-status-row" });
  const statusBadges = createStatusBadges(deliverable);
  const statusDropdown = createStatusDropdown(deliverable, project);
  const detailsToggle = createDetailsToggle(card);
  statusSection.append(statusBadges, statusDropdown, detailsToggle);

  // Tasks preview (2-3 tasks, now clickable)
  const tasksPreview = createTasksPreview(deliverable, card);

  // Notes section (collapsible)
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
    if (p.nick)
      nameCell.append(
        el("small", { className: "muted", textContent: ` (${p.nick})` })
      );
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
  const data = readForm();
  autoSetPrimary(data);
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

  if (insertAtTop && list.firstChild) {
    list.prepend(card);
  } else {
    list.appendChild(card);
  }

  const hasPrimary = list.querySelector(".d-primary:checked");
  if (!hasPrimary && !primaryId) primaryInput.checked = true;
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
  row.querySelector(".t-text").value = t.text || "";
  row.querySelector(".t-done").checked = !!t.done;
  row.querySelector(".t-link").value = t.links?.[0]?.raw || "";
  row.querySelector(".t-link2").value = t.links?.[1]?.raw || "";
  row.querySelector(".btn-remove").onclick = () => row.remove();
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

window.addDeliverable = () =>
  addDeliverableCard(createDeliverable(), null, { insertAtTop: true });
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

async function loadAndRenderBundles() {
  const container = document.getElementById("commands-container");
  if (!container) return;
  container.innerHTML = '<div class="spinner">Loading...</div>';

  try {
    const response = await window.pywebview.api.get_bundle_statuses();
    if (response.status !== "success") throw new Error(response.message);

    container.innerHTML = "";
    if (response.data.length === 0) {
      container.textContent = "No command bundles found.";
      return;
    }

    // Process each bundle
    for (const bundle of response.data) {
      // Normalize name
      const coreName = bundle.name
        .replace("ElectricalCommands.", "")
        .replace(".bundle", "");

      // Fetch description (background)
      const description = await fetchDescriptionForBundle(bundle.name);

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
    }
  } catch (e) {
    container.innerHTML = `<div class="error-message">Error: ${e.message}</div>`;
  }
}

function prefetchBundles() {
  if (bundlesPrefetchStarted) return;
  bundlesPrefetchStarted = true;
  setTimeout(() => {
    loadAndRenderBundles();
  }, 0);
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
    await loadAndRenderBundles();

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
      loadAndRenderBundles();
    } else if (tab === "projects") {
      render();
    } else if (tab === "timesheets") {
      renderTimesheets();
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
        const baseKey = getProjectBaseKey(aiProject?.path);
        const matchIndex = baseKey
          ? db.findIndex((project) => getProjectBaseKey(project?.path) === baseKey)
          : -1;
        if (matchIndex >= 0) {
          const target = normalizeProject(db[matchIndex]);
          const incomingDeliverables = Array.isArray(aiProject?.deliverables)
            ? aiProject.deliverables.map(normalizeDeliverable)
            : [createDeliverable()];
          target.deliverables.push(...incomingDeliverables);
          autoSetPrimary(target);
          db[matchIndex] = target;
          save();
          render();
          toast("Added deliverable to existing project.");
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
        loadAndRenderBundles();
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
        loadAndRenderBundles();
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
      const [loadedDb, loadedNotes, loadedTimesheets] = await Promise.all([
        load(),
        loadNotes(),
        loadTimesheets(),
      ]);
      db = loadedDb;
      notesDb = loadedNotes || {};
      timesheetDb = loadedTimesheets || { weeks: {}, lastModified: null };
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
