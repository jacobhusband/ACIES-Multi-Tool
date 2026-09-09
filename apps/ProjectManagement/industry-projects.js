// Industry presentation layer for the Projects tab.
//
// Implements the "Deliverables UI" design: the register row (list view), the
// deliverable plate (board view), one global command line docked to the bottom
// of the application.
//
// Everything here reads the same project + deliverable records script.js owns
// and calls back into its existing handlers (status changes, calendar picker,
// tool launches, notes, folders) so behaviour stays identical — only the
// presentation changes. Loaded after script.js.

const INDUSTRY_STATUS_OPTIONS = Object.freeze([
  "Waiting",
  "On hold",
  "Pending Review",
  "Complete",
  "Completed (by others)",
  "Delivered",
]);

const INDUSTRY_FREQUENT_TOOL_IDS = Object.freeze([
  "toolCleanDrawings",
  "toolPublishDwgs",
  "toolCreatePlanCheckTemplate",
]);


// ---------------------------------------------------------------------------
// Date + state helpers
// ---------------------------------------------------------------------------

function industryToday() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return today;
}

function industryDayDiff(date) {
  if (!(date instanceof Date) || isNaN(date)) return null;
  const target = new Date(date);
  target.setHours(0, 0, 0, 0);
  return Math.round((target - industryToday()) / 86400000);
}

function industryPlural(count, singular, plural = `${singular}s`) {
  return `${count} ${count === 1 ? singular : plural}`;
}

function industryRelativeDays(days) {
  if (days === null || days === undefined) return "";
  if (days === 0) return "today";
  if (days > 0) return `in ${industryPlural(days, "day")}`;
  return `${industryPlural(-days, "day")} late`;
}

function getDeliverablePrimaryStatus(deliverable) {
  return INDUSTRY_STATUS_OPTIONS.find((status) => hasStatus(deliverable, status)) || "";
}

function getProjectClientLabel(project) {
  const parts = [];
  const account = extractAccountFromPath(project?.path);
  if (account) parts.push(account);
  const nick = String(project?.nick || "").trim();
  if (nick && nick !== account) parts.push(nick);
  return parts.join(" · ");
}

function getProjectShortName(project) {
  const nick = String(project?.nick || "").trim();
  if (nick) return nick;
  const name = String(project?.name || "").trim();
  if (!name) return String(project?.id || "").trim();
  const first = name.split(",")[0].trim();
  return first.length > 28 ? `${first.slice(0, 27)}…` : first;
}

// One read of a deliverable that every Industry surface shares:
//   kind  — late | nostatus | open | done
//   rail  — hatch | dashed | accent | muted   (the left rail / bar treatment)
//   stamp — the status object drawn in the Status column
function getDeliverableRegisterState(deliverable) {
  const finished = isFinished(deliverable);
  const status = getDeliverablePrimaryStatus(deliverable);
  const externalDate = parseDueStr(getHardDueStr(deliverable));
  const effectiveDate = parseDueStr(getEffectiveDueStr(deliverable));
  const externalDays = externalDate ? industryDayDiff(externalDate) : null;
  const effectiveDays = effectiveDate ? industryDayDiff(effectiveDate) : null;

  if (finished) {
    return {
      kind: "done",
      rail: "muted",
      status,
      stamp: { text: status || "Complete", style: "neutral" },
      sub: "",
      daysLate: 0,
      daysAhead: effectiveDays,
      externalLate: false,
    };
  }

  const externalLate = externalDays !== null && externalDays < 0;
  const internalLate = effectiveDays !== null && effectiveDays < 0;
  if (externalLate || internalLate) {
    const daysLate = externalLate ? -externalDays : -effectiveDays;
    return {
      kind: "late",
      rail: "hatch",
      status,
      stamp: {
        text: `${daysLate} ${daysLate === 1 ? "DAY" : "DAYS"} LATE`,
        style: "solid",
      },
      sub: status || "No status",
      daysLate,
      daysAhead: effectiveDays,
      externalLate,
    };
  }

  const relative = effectiveDays === null ? "" : industryRelativeDays(effectiveDays);
  if (!status) {
    return {
      kind: "nostatus",
      rail: "dashed",
      status,
      stamp: { text: "NO STATUS", style: "dashed" },
      sub: effectiveDays === null ? "no date set" : `due ${relative}`,
      daysLate: 0,
      daysAhead: effectiveDays,
      externalLate: false,
    };
  }

  return {
    kind: "open",
    rail: "accent",
    status,
    stamp: { text: status, style: "tag" },
    sub: effectiveDays === null ? "no date set" : relative,
    daysLate: 0,
    daysAhead: effectiveDays,
    externalLate: false,
  };
}

function countDeliverableNotes(deliverable) {
  const noteItems = Array.isArray(deliverable?.noteItems) ? deliverable.noteItems : [];
  return noteItems.filter((item) => String(item?.text || item || "").trim()).length;
}

function getAllProjectDeliverableRows() {
  const rows = [];
  (Array.isArray(db) ? db : []).forEach((project) => {
    const deliverables =
      typeof getOverviewDeliverables === "function"
        ? getOverviewDeliverables(project)
        : Array.isArray(project?.deliverables)
          ? project.deliverables
          : [];
    deliverables.forEach((deliverable) => {
      if (!deliverable) return;
      rows.push({ project, deliverable });
    });
  });
  return rows;
}

// Clicks on controls inside a row must not count as "select this row".
function isInteractiveTarget(target) {
  return !!target?.closest?.(
    "button, a, input, textarea, select, [role='button'], [contenteditable='true']"
  );
}

// ---------------------------------------------------------------------------
// Shared pieces: stamps, date fields, action stamps
// ---------------------------------------------------------------------------

function createRegisterStamp(state, deliverable, project) {
  const stamp = el("button", {
    type: "button",
    className: `reg-stamp reg-stamp--${state.stamp.style}`,
    textContent: state.stamp.text,
    title: "Select this deliverable and choose a status in the command line",
    "aria-label": `Status: ${state.stamp.text}. Select to change status.`,
  });
  stamp.addEventListener("click", (event) => {
    event.preventDefault();
    event.stopPropagation();
    selectDeliverableForCommands(deliverable, project, { expand: true, focus: "status" });
  });
  return stamp;
}

function createRegisterDateField(deliverable, project, field, state) {
  const value =
    field === "hardDue"
      ? getHardDueStr(deliverable)
      : String(deliverable?.due || "").trim();
  const label = field === "hardDue" ? "External" : "Internal";
  const wrap = el("div", {
    className: `reg-field reg-field--${field === "hardDue" ? "external" : "internal"}`,
    role: "button",
    tabIndex: 0,
    title: value
      ? `${label} date ${humanDate(value)}. Click to change.`
      : `Set the ${label.toLowerCase()} date.`,
  });
  const kick = el("div", { className: "reg-kick", textContent: label });
  const val = el("div", {
    className: `reg-val${value ? "" : " is-empty"}`,
    textContent: value ? humanDate(value) : "not set",
  });
  if (field === "hardDue" && value) {
    wrap.classList.add("is-external");
    if (state.kind === "late" && state.externalLate) wrap.classList.add("is-late");
  }
  const open = (event) => {
    event.preventDefault();
    event.stopPropagation();
    showCalendarForDeliverableBadge(wrap, deliverable, project, field);
  };
  wrap.addEventListener("click", open);
  wrap.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") open(event);
  });
  wrap.append(kick, val);
  return wrap;
}

function createRegisterIco({ text, title, className = "", onClick }) {
  const button = el("button", {
    type: "button",
    className: `reg-ico ${className}`.trim(),
    textContent: text,
    title,
    "aria-label": title,
  });
  button.draggable = false;
  button.addEventListener("dragstart", (event) => {
    event.preventDefault();
    event.stopPropagation();
  });
  button.addEventListener("click", (event) => {
    event.preventDefault();
    event.stopPropagation();
    const result = onClick?.(event);
    if (result && typeof result.catch === "function") {
      result.catch((error) => console.warn("Deliverable action failed:", error));
    }
  });
  return button;
}

// Notes stay in one button; the corner badge counts important project notes.
function addImportantNotesBadge(button, project) {
  const count = getProjectImportantItems(project).length;
  button.classList.add("notes-badge-host");
  if (!count) return;
  button.classList.add("has-count");
  const label = `${button.title}. ${industryPlural(count, "important note")}`;
  button.title = label;
  button.setAttribute("aria-label", label);
  button.appendChild(el("span", {
    className: "notes-important-badge",
    textContent: String(count),
    "aria-hidden": "true",
  }));
}

function createRegisterActions(deliverable, project) {
  const wrap = el("div", { className: "reg-actions" });
  if (!project) return wrap;
  const noteCount = countDeliverableNotes(deliverable);
  const notes = createRegisterIco({
    text: noteCount ? `N${noteCount}` : "N",
    className: noteCount ? "has-count" : "",
    title: noteCount
      ? `${industryPlural(noteCount, "note")} — open project notes`
      : "Open project notes",
    onClick: () => openProjectPage(project),
  });
  addImportantNotesBadge(notes, project);
  wrap.appendChild(notes);
  return wrap;
}

function attachDeliverableRename(nameEl, deliverable) {
  nameEl.title = "Double-click to rename";
  nameEl.addEventListener("dblclick", (event) => {
    event.preventDefault();
    event.stopPropagation();
    const input = el("input", {
      className: "reg-code-input",
      type: "text",
      value: deliverable.name || "",
    });
    nameEl.hidden = true;
    nameEl.parentNode.insertBefore(input, nameEl.nextSibling);
    input.focus();
    input.select();
    let finished = false;
    const finish = async (shouldSave) => {
      if (finished) return;
      finished = true;
      const next = input.value.trim();
      if (shouldSave && next && next !== deliverable.name) {
        deliverable.name = next;
        nameEl.textContent = next;
        await save();
      }
      input.remove();
      nameEl.hidden = false;
    };
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        finish(true);
      } else if (e.key === "Escape") {
        e.preventDefault();
        finish(false);
      }
    });
    input.addEventListener("blur", () => finish(true));
    input.addEventListener("click", (e) => e.stopPropagation());
  });
}

// ---------------------------------------------------------------------------
// 1a — Register row (list view)
// ---------------------------------------------------------------------------

function buildDeliverableRegisterColumns(deliverable, project) {
  syncDeliverableWorkItemFields(deliverable);
  if (!deliverable.id) deliverable.id = createId("dlv");
  const state = getDeliverableRegisterState(deliverable);

  const code = el("div", {
    className: "reg-code",
    textContent: deliverable.name || "Deliverable",
  });
  attachDeliverableRename(code, deliverable);

  const internal = createRegisterDateField(deliverable, project, "due", state);
  const external = createRegisterDateField(deliverable, project, "hardDue", state);

  const statusCol = el("div", { className: "reg-status" });
  statusCol.appendChild(createRegisterStamp(state, deliverable, project));
  if (state.sub) {
    statusCol.appendChild(el("div", { className: "reg-sub", textContent: state.sub }));
  }

  const actions = createRegisterActions(deliverable, project);
  return [code, internal, external, statusCol, actions];
}

function renderDeliverableRegisterCell(cell, deliverable, project) {
  if (!cell) return;
  cell.innerHTML = "";
  cell.classList.add("reg-cell");
  if (!deliverable) {
    cell.appendChild(el("div", { className: "reg-code is-empty", textContent: "—" }));
    return;
  }
  buildDeliverableRegisterColumns(deliverable, project).forEach((node) =>
    cell.appendChild(node)
  );
}

// Grouped-by-project mode stacks every visible deliverable of a project in
// one row; each becomes its own five-column line under the shared project.
function renderDeliverableRegisterCellGroup(cell, deliverables, project, note = "") {
  if (!cell) return;
  cell.innerHTML = "";
  cell.classList.add("reg-cell", "reg-cell--multi");
  const stack = el("div", { className: "reg-multi" });
  if (note) {
    stack.appendChild(el("div", { className: "reg-note", textContent: note }));
  }
  if (!deliverables.length) {
    stack.appendChild(el("div", { className: "reg-code is-empty", textContent: "—" }));
  }
  deliverables.forEach((deliverable) => {
    const line = el("div", {
      className: `reg-line reg-line--${getDeliverableRegisterState(deliverable).kind}`,
      tabIndex: 0,
    });
    line.dataset.deliverableId = String(deliverable.id || "");
    buildDeliverableRegisterColumns(deliverable, project).forEach((node) =>
      line.appendChild(node)
    );
    attachDeliverableSelection(line, deliverable, project);
    stack.appendChild(line);
  });
  cell.appendChild(stack);
}

// Clicking (or pressing Enter on) a row makes it the target of the command line.
function attachDeliverableSelection(element, deliverable, project) {
  element.addEventListener("click", (event) => {
    if (isInteractiveTarget(event.target)) return;
    if (window.getSelection?.()?.toString()) return;
    selectDeliverableForCommands(deliverable, project);
  });
  element.addEventListener("keydown", (event) => {
    if (event.target !== element) return;
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      selectDeliverableForCommands(deliverable, project, { expand: true });
    }
  });
}

// Turns the plain table row script.js builds into a register row: adds the
// rail, moves the account/nick under the project name, marks the state.
function decorateProjectRegisterRow(tr, project, deliverable) {
  if (!tr) return;
  const state = deliverable ? getDeliverableRegisterState(deliverable) : null;
  tr.classList.add("reg-row");
  if (state) tr.classList.add(`reg-row--${state.kind}`);
  if (deliverable && isDeliverablePinned(deliverable)) tr.classList.add("is-pinned");
  if (deliverable?.id) tr.dataset.deliverableId = String(deliverable.id);

  const idCell = tr.querySelector(".cell-id");
  if (idCell && !idCell.querySelector(".reg-rail")) {
    idCell.insertBefore(
      el("span", { className: `reg-rail reg-rail--${state?.rail || "muted"}` }),
      idCell.firstChild
    );
  }

  const main = tr.querySelector(".project-details-main");
  if (main) {
    const smalls = [...main.querySelectorAll("small.muted")];
    const clientText = smalls
      .map((node) => node.textContent.replace(/^\s*\(|\)\s*$/g, "").trim())
      .filter(Boolean)
      .join(" · ");
    smalls.forEach((node) => node.remove());
    main.querySelector(".path-link, .project-title-text")?.classList.add("reg-pname");
    if (clientText) {
      main.appendChild(el("div", { className: "reg-client", textContent: clientText }));
    }
  }

  if (deliverable && !groupDeliverablesByProject) {
    tr.tabIndex = 0;
    attachDeliverableSelection(tr, deliverable, project);
  }
}

// Numbers each section divider ("Pinned — 01") and fills the footer count.
function finalizeProjectsRegister(tbody, pagination) {
  if (!tbody) return;
  let sectionIndex = 0;
  [...tbody.children].forEach((row) => {
    if (!row.classList.contains("week-separator-row")) return;
    sectionIndex += 1;
    const separator = row.querySelector(".week-separator");
    if (!separator) return;
    let count = separator.querySelector(".reg-sep-count");
    if (!count) {
      count = el("span", { className: "reg-sep-count" });
      separator.appendChild(count);
    }
    count.textContent = String(sectionIndex).padStart(2, "0");
  });

  const footer = document.getElementById("registerFooterCount");
  if (footer) {
    const shown = Array.isArray(pagination?.items) ? pagination.items.length : 0;
    const total = getAllProjectDeliverableRows().length;
    footer.textContent = `${shown} of ${total} shown`;
  }
  syncCommandDockSelection();
}

// ---------------------------------------------------------------------------
// 1b — Deliverable plate (board view)
// ---------------------------------------------------------------------------

function createBlueprintCorners() {
  return ["tl", "tr", "bl", "br"].map((pos) => el("i", { className: `corner ${pos}` }));
}

function createPlateRuleCell(label, value, { muted = false, filled = false, strong = false, hatched = false, onClick = null, title = "" } = {}) {
  const cell = el("div", {
    className: `plate-rule-cell${filled ? " is-filled" : ""}${hatched ? " is-hatched" : ""}`,
  });
  cell.append(
    el("div", { className: "reg-kick", textContent: label }),
    el("div", {
      className: `reg-val${muted ? " is-empty" : ""}${strong ? " is-strong" : ""}`,
      textContent: value,
    })
  );
  if (onClick) {
    cell.classList.add("is-clickable");
    cell.setAttribute("role", "button");
    cell.tabIndex = 0;
    if (title) cell.title = title;
    cell.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      onClick(event);
    });
    cell.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        onClick(event);
      }
    });
  }
  return cell;
}

function renderDeliverablePlateCard(deliverable, project) {
  syncDeliverableWorkItemFields(deliverable);
  const deliverableId = String(deliverable?.id || createId("dlv")).trim();
  if (!deliverable.id) deliverable.id = deliverableId;
  const state = getDeliverableRegisterState(deliverable);

  const card = el("div", {
    className: `deliverable-card-new plate-card blueprint plate-card--${state.kind}${
      isDeliverablePinned(deliverable) ? " is-pinned-deliverable" : ""
    }`,
    tabIndex: 0,
  });
  card.dataset.deliverableId = deliverableId;
  createBlueprintCorners().forEach((corner) => card.appendChild(corner));

  // Head: project id · client / deliverable title / stamp
  const head = el("div", { className: "plate-head" });
  const identity = el("div", { className: "plate-identity" });
  const kickParts = [String(project?.id || "").trim(), getProjectClientLabel(project)].filter(Boolean);
  identity.appendChild(el("div", { className: "reg-kick plate-kick", textContent: kickParts.join(" · ") || "—" }));
  const title = el("div", { className: "plate-title", textContent: deliverable.name || "Deliverable" });
  attachDeliverableRename(title, deliverable);
  identity.appendChild(title);
  const stamp = createRegisterStamp(state, deliverable, project);
  if (state.kind === "open" && state.daysAhead !== null) {
    stamp.textContent = industryRelativeDays(state.daysAhead);
    stamp.setAttribute("aria-label", `${state.status}, ${stamp.textContent}. Select to change status.`);
  }
  head.append(identity, stamp);
  card.appendChild(head);

  // Project name (opens folder / context menu exactly like the list)
  const projectName = String(project?.name || "").trim();
  if (projectName) {
    const nameEl = el("div", { className: "plate-project", textContent: projectName, title: projectName, role: "button" });
    attachProjectDirectoryContextMenu(nameEl, project);
    attachCardProjectPathOpen(nameEl, project);
    card.appendChild(nameEl);
  }

  // Milestone rule: internal → external → delivered
  const rule = el("div", { className: "plate-rule" });
  const internalValue = String(deliverable?.due || "").trim();
  const externalValue = getHardDueStr(deliverable);
  rule.appendChild(
    createPlateRuleCell("Internal", internalValue ? humanDate(internalValue) : "not set", {
      muted: !internalValue,
      title: "Click to change the internal date",
      onClick: (event) => {
        showCalendarForDeliverableBadge(event.currentTarget, deliverable, project, "due");
      },
    })
  );
  rule.appendChild(
    createPlateRuleCell("External", externalValue ? humanDate(externalValue) : "not set", {
      muted: !externalValue,
      strong: !!externalValue && state.kind !== "done",
      hatched: state.kind === "late" && state.externalLate,
      title: "Click to change the external date",
      onClick: (event) => {
        showCalendarForDeliverableBadge(event.currentTarget, deliverable, project, "hardDue");
      },
    })
  );
  rule.appendChild(
    createPlateRuleCell("Delivered", state.kind === "done" ? state.status || "Complete" : "—", {
      muted: state.kind !== "done",
      filled: state.kind === "done",
    })
  );
  card.appendChild(rule);

  // Foot
  const foot = el("div", { className: "plate-foot" });
  if (state.kind === "done") {
    foot.appendChild(el("span", { className: "reg-kick", textContent: "Closed" }));
  } else if (state.status) {
    const tag = el("button", {
      type: "button",
      className: "reg-stamp reg-stamp--outline",
      textContent: state.status,
      title: "Select this deliverable and choose a status in the command line",
    });
    tag.addEventListener("click", (event) => {
      event.stopPropagation();
      selectDeliverableForCommands(deliverable, project, { expand: true, focus: "status" });
    });
    foot.appendChild(tag);
  }
  foot.appendChild(el("span", { className: "plate-spacer" }));
  if (project) {
    const notes = el("button", {
      type: "button",
      className: "plate-notes reg-kick",
      textContent: "NOTES",
      title: "Open project notes",
    });
    notes.addEventListener("click", (event) => {
      event.stopPropagation();
      openProjectPage(project);
    });
    addImportantNotesBadge(notes, project);
    foot.appendChild(notes);
  }
  card.appendChild(foot);

  attachDeliverableSelection(card, deliverable, project);
  return card;
}

// ---------------------------------------------------------------------------
// 1c — Command dock: one command line for the whole application
// ---------------------------------------------------------------------------
//
// Works in either order: type an action and then pick a deliverable, or pick a
// deliverable and then type. The line also searches: projects first, then the
// chosen project's deliverables, then commands. The action list stays
// collapsed under the line until the user expands it.

let industryCommandDock = null;

function ensureCommandDock() {
  if (industryCommandDock) return industryCommandDock;

  const root = el("div", {
    className: "cmd-dock",
    id: "commandDock",
    role: "region",
    "aria-label": "Command line",
  });
  root.dataset.expanded = "false";

  // Expandable action list (above the line)
  const body = el("div", { className: "cmd-dock-body" });
  body.hidden = true;
  const columns = el("div", { className: "cmd-body" });
  const left = el("div", { className: "cmd-col cmd-col--left" });
  const right = el("div", { className: "cmd-col cmd-col--right" });
  columns.append(left, right);
  const foot = el("div", { className: "cmd-foot" });
  const scope = el("span", { className: "reg-kick cmd-scope" });
  foot.append(
    el("span", { className: "reg-kick", textContent: "↑↓ move · ↩ run · esc clear / collapse" }),
    el("span", { className: "cmd-spacer" }),
    scope
  );
  body.append(columns, foot);

  // The line itself
  const bar = el("div", { className: "cmd-dock-bar" });
  const context = el("button", {
    type: "button",
    className: "reg-kick cmd-context",
    title: "Selected deliverable — click to clear",
  });
  const input = el("input", {
    type: "text",
    className: "cmd-search",
    placeholder: "Type a command…",
    "aria-label": "Command line",
    autocomplete: "off",
    spellcheck: false,
  });
  const pending = el("span", { className: "reg-kick cmd-pending" });
  pending.hidden = true;
  const hint = el("span", { className: "reg-kick cmd-key", textContent: "Ctrl K" });
  const expand = el("button", {
    type: "button",
    className: "cmd-expand",
    "aria-expanded": "false",
    title: "Show available actions",
  });
  expand.append(
    el("span", { className: "reg-kick", textContent: "Actions" }),
    el("span", { className: "cmd-expand-chevron", "aria-hidden": "true", textContent: "▴" })
  );
  bar.append(context, input, pending, hint, expand);

  root.append(body, bar);
  document.body.appendChild(root);

  industryCommandDock = {
    root,
    body,
    left,
    right,
    scope,
    context,
    input,
    pending,
    expand,
    deliverable: null,
    project: null,
    pendingKey: null,
    pendingLabel: "",
    items: [],
    sections: [],
    activeIndex: -1,
    expanded: false,
  };

  input.addEventListener("input", () => {
    if (input.value.trim() && !industryCommandDock.expanded) setCommandDockExpanded(true);
    filterCommandDock();
  });
  input.addEventListener("keydown", handleCommandDockKeydown);
  // Capture before outside controls run: explicit status/action triggers can
  // still reopen the sheet, while ordinary outside clicks simply dismiss it.
  document.addEventListener("pointerdown", (event) => {
    if (industryCommandDock.expanded && !root.contains(event.target)) {
      setCommandDockExpanded(false);
    }
  }, true);
  body.addEventListener("keydown", (event) => {
    if (event.target !== input) handleCommandDockKeydown(event);
  });
  expand.addEventListener("click", () => {
    setCommandDockExpanded(!industryCommandDock.expanded);
    if (industryCommandDock.expanded) input.focus({ preventScroll: true });
  });
  context.addEventListener("click", () => {
    if (industryCommandDock.deliverable || industryCommandDock.project) {
      clearCommandDockSelection();
    } else {
      setCommandDockExpanded(true);
    }
    input.focus({ preventScroll: true });
  });

  renderCommandDockItems();
  updateCommandDockContext();
  return industryCommandDock;
}

function setCommandDockExpanded(expanded) {
  const dock = ensureCommandDock();
  dock.expanded = !!expanded;
  dock.body.hidden = !dock.expanded;
  dock.root.dataset.expanded = String(dock.expanded);
  dock.expand.setAttribute("aria-expanded", String(dock.expanded));
  dock.expand.title = dock.expanded ? "Hide available actions" : "Show available actions";
  if (dock.expanded) filterCommandDock();
}

function getCommandDockProjectDeliverables(project) {
  if (!project) return [];
  const list =
    typeof getOverviewDeliverables === "function"
      ? getOverviewDeliverables(project)
      : Array.isArray(project.deliverables)
        ? project.deliverables
        : [];
  const sorted = list.filter(Boolean).slice();
  if (typeof compareDeliverablesByDue === "function") sorted.sort(compareDeliverablesByDue);
  return sorted;
}

function updateCommandDockContext() {
  const dock = ensureCommandDock();
  const { deliverable, project } = dock;
  dock.context.classList.remove("has-selection", "has-project");
  if (deliverable) {
    const parts = [
      String(project?.id || "").trim(),
      String(deliverable?.name || "Deliverable").trim(),
      getProjectShortName(project),
    ].filter(Boolean);
    dock.context.textContent = parts.join(" · ");
    dock.context.classList.add("has-selection");
    dock.context.title = "Selected deliverable — click to go back to the project";
    dock.scope.textContent = "Acts on 1 deliverable";
    dock.input.placeholder = "Type a command…";
  } else if (project) {
    const parts = [String(project?.id || "").trim(), getProjectShortName(project)].filter(Boolean);
    dock.context.textContent = `${parts.join(" · ")} · choose a deliverable`;
    dock.context.classList.add("has-project");
    dock.context.title = "Selected project — click to clear";
    dock.scope.textContent = "Type to find one of this project's deliverables";
    dock.input.placeholder = dock.pendingKey
      ? "Now choose a deliverable…"
      : "Search this project's deliverables, or type a command…";
  } else {
    dock.context.textContent = "No project selected";
    dock.context.title = "Type a project number or name, or click a row";
    dock.scope.textContent = "Search a project, then a deliverable, then run a command";
    dock.input.placeholder = dock.pendingKey
      ? "Now search a project or select a row…"
      : "Search projects or type a command…";
  }
  if (dock.pendingKey) {
    dock.pending.textContent = `${dock.pendingLabel} · waiting for a deliverable`;
    dock.pending.hidden = false;
  } else {
    dock.pending.hidden = true;
  }
  dock.root.classList.toggle("has-selection", !!deliverable);
  dock.root.classList.toggle("has-project", !!project && !deliverable);
  dock.root.classList.toggle("has-pending", !!dock.pendingKey);
}

// The "Target" group: projects until one is chosen, then that project's
// deliverables until one is chosen. Typing in the line searches it.
function buildCommandDockTargetGroup(deliverable, project) {
  if (deliverable) return null;
  if (project) {
    const items = getCommandDockProjectDeliverables(project).map((candidate) => {
      const state = getDeliverableRegisterState(candidate);
      const dateStr = String(candidate?.due || "").trim() || getHardDueStr(candidate);
      const meta = [state.status || (state.kind === "done" ? "" : "no status"), dateStr ? humanDate(dateStr) : "", state.kind === "late" ? state.stamp.text.toLowerCase() : state.sub]
        .filter(Boolean)
        .join(" · ");
      return {
        key: `deliverable:${candidate.id}`,
        kind: "deliverable",
        label: String(candidate?.name || "Deliverable"),
        search: `${candidate?.name || ""} ${meta}`,
        meta,
        railKind: state.kind,
        run: () => selectDeliverableForCommands(candidate, project, { expand: true }),
      };
    });
    return {
      key: "target",
      column: "left",
      label: `Deliverables · ${getProjectShortName(project)}`,
      items,
      limit: 12,
      empty: "This project has no deliverables.",
    };
  }
  const items = (Array.isArray(db) ? db : []).map((candidate) => {
    const id = String(candidate?.id || "").trim();
    const name = String(candidate?.name || "").trim();
    const client = getProjectClientLabel(candidate);
    const count = getCommandDockProjectDeliverables(candidate).length;
    return {
      key: `project:${id || name}`,
      kind: "project",
      label: name || id || "Project",
      prefix: id,
      search: `${id} ${name} ${client} ${candidate?.nick || ""}`,
      meta: [client, industryPlural(count, "deliverable")].filter(Boolean).join(" · "),
      run: () => selectProjectForCommands(candidate),
    };
  });
  return {
    key: "target",
    column: "left",
    label: "Projects",
    items,
    limit: 10,
    empty: "No projects yet.",
  };
}

// Every command, keyed so a pending action survives selection and re-renders.
function buildCommandDockGroups(deliverable, project) {
  const hasTarget = !!deliverable;
  const rerender = () => renderProjectsPreservingExpandedDeliverables();

  const statusItems = INDUSTRY_STATUS_OPTIONS.map((status) => ({
    key: `status:${status}`,
    label: status,
    kind: "status",
    checked: hasTarget && hasStatus(deliverable, status),
    run: async (target) => {
      setSingleStatus(target.deliverable, status);
      await save();
      rerender();
    },
  }));
  statusItems.push({
    key: "status:clear",
    label: "Clear status",
    kind: "status-clear",
    hidden: hasTarget && !getDeliverablePrimaryStatus(deliverable),
    run: async (target) => {
      setSingleStatus(target.deliverable, "");
      await save();
      rerender();
    },
  });

  const deliverableItems = [
    {
      key: "pin",
      label: hasTarget && isDeliverablePinned(deliverable) ? "Unpin deliverable" : "Pin deliverable",
      run: async (target) => {
        setDeliverablePinnedState(target.deliverable, !isDeliverablePinned(target.deliverable));
        await save();
        rerender();
      },
    },
    {
      key: "notes",
      label: "Open project notes",
      run: (target) => openProjectPage(target.project),
    },
    {
      key: "folder",
      label: "Open project folder",
      hidden: hasTarget && !project?.path,
      run: async (target) => {
        if (!target.project?.path) {
          toast("This project has no folder path.");
          return;
        }
        if (!window.pywebview?.api?.open_path) {
          toast("Open path is unavailable.");
          return;
        }
        try {
          const result = await window.pywebview.api.open_path(convertPath(target.project.path));
          if (result?.status && result.status !== "success") {
            throw new Error(result.message || "Unable to open folder.");
          }
          toast("Opening folder...");
        } catch (error) {
          toast(error?.message || "Failed to open path.");
        }
      },
    },
    {
      key: "attachments",
      label: "Attachments",
      run: (target) =>
        openDeliverableAttachmentsFromDock(
          target.deliverable,
          target.project,
          findDeliverableElement(target.deliverable) || industryCommandDock.context
        ),
    },
    {
      key: "edit-project",
      label: "Edit project",
      run: (target) => {
        const index = Array.isArray(db) ? db.indexOf(target.project) : -1;
        if (index >= 0) openEdit(index);
      },
    },
    {
      key: "delete",
      label: "Delete deliverable",
      danger: true,
      run: (target) => removeDeliverable(target.project, target.deliverable),
    },
  ];

  const toolEntries = getDeliverableToolMenuEntries();
  const frequent = [];
  INDUSTRY_FREQUENT_TOOL_IDS.forEach((toolId, index) => {
    const entry = toolEntries.find((candidate) => candidate.id === toolId);
    if (!entry) return;
    frequent.push({
      key: `tool:${entry.id}`,
      label: entry.menuLabel || entry.label,
      hintKey: `Ctrl ${index + 1}`,
      shortcut: String(index + 1),
      run: (target) =>
        launchSharedToolCard(
          entry.id,
          buildProjectsTabToolLaunchContext(target.project, target.deliverable)
        ),
    });
  });
  DELIVERABLE_QUICK_ACCESS_ACTIONS.forEach((action, index) => {
    frequent.push({
      key: `quick:${action.id}`,
      label: getDeliverableQuickAccessActionLabel(action),
      hintKey: index === 0 ? "Ctrl ⏎" : "",
      shortcut: index === 0 ? "Enter" : "",
      run: (target) => void action.run(target.project, target.deliverable),
    });
  });
  const allTools = toolEntries
    .filter((entry) => !INDUSTRY_FREQUENT_TOOL_IDS.includes(entry.id))
    .map((entry) => ({
      key: `tool:${entry.id}`,
      label: entry.menuLabel || entry.label,
      tag: true,
      run: (target) =>
        launchSharedToolCard(
          entry.id,
          buildProjectsTabToolLaunchContext(target.project, target.deliverable)
        ),
    }));

  const targetGroup = buildCommandDockTargetGroup(deliverable, project);
  return [
    ...(targetGroup ? [targetGroup] : []),
    { key: "status", column: "left", label: "Status", items: statusItems },
    { key: "deliverable", column: "left", label: "Deliverable", items: deliverableItems },
    { key: "tools-frequent", column: "right", label: "Tools · frequent", items: frequent },
    { key: "tools-all", column: "right", label: "Tools · all", items: allTools, tags: true },
  ];
}

function findDeliverableElement(deliverable) {
  const id = String(deliverable?.id || "").trim();
  if (!id) return null;
  const selector = `[data-deliverable-id="${CSS.escape(id)}"]`;
  return (
    document.querySelector(`#projectsCardView ${selector}`) ||
    document.querySelector(`#tbody ${selector}`)
  );
}

function openDeliverableAttachmentsFromDock(deliverable, project, anchor) {
  ensureAttachmentPanel();
  const descriptor = {
    kind: "deliverable",
    owner: deliverable,
    deliverable,
    project,
    scope: "projects-tab",
  };
  const context = {
    ...descriptor,
    trigger: anchor || document.body,
    getAttachments: () => getAttachmentOwnerAttachments(descriptor),
    setAttachments: async (next) =>
      setAttachmentOwnerAttachments(descriptor, next, {
        persistNow: true,
        onChange: () => renderProjectsPreservingExpandedDeliverables(),
      }),
  };
  openAttachmentPanel(context);
}

function renderCommandDockItems() {
  const dock = industryCommandDock;
  const groups = buildCommandDockGroups(dock.deliverable, dock.project);
  dock.left.innerHTML = "";
  dock.right.innerHTML = "";
  dock.items = [];
  dock.sections = [];
  dock.activeIndex = -1;

  groups.forEach((group) => {
    const host = group.column === "left" ? dock.left : dock.right;
    const section = el("div", { className: `cmd-group cmd-group--${group.key}` });
    section.appendChild(el("div", { className: "reg-kick cmd-group-label", textContent: group.label }));
    const list = el("div", { className: `cmd-list${group.tags ? " cmd-list--tags" : ""}` });
    group.items.forEach((item) => {
      if (item.hidden) return;
      const isTarget = item.kind === "project" || item.kind === "deliverable";
      const node = el("button", {
        type: "button",
        className: `cmd-item${item.tag ? " cmd-item--tag" : ""}${item.danger ? " is-danger" : ""}${
          item.kind === "status" ? " cmd-item--status" : ""
        }${isTarget ? ` cmd-item--target cmd-item--${item.kind}` : ""}${item.checked ? " is-checked" : ""}`,
      });
      node.setAttribute("role", "menuitem");
      node.dataset.commandKey = item.key;
      if (item.kind === "status") {
        node.appendChild(el("span", { className: "cmd-radio", "aria-hidden": "true" }));
      }
      if (item.kind === "deliverable") {
        node.appendChild(el("span", { className: `reg-rail reg-rail--${getRailForKind(item.railKind)} cmd-item-rail` }));
      }
      if (item.prefix) {
        node.appendChild(el("span", { className: "reg-kick cmd-item-prefix", textContent: item.prefix }));
      }
      node.appendChild(el("span", { className: "cmd-item-label", textContent: item.label }));
      if (item.meta) {
        node.appendChild(el("span", { className: "cmd-item-meta", textContent: item.meta }));
      }
      if (item.hintKey) {
        node.appendChild(el("span", { className: "reg-kick cmd-key-hint", textContent: item.hintKey }));
      }
      node.addEventListener("click", (event) => {
        event.preventDefault();
        event.stopPropagation();
        runCommandDockItem(item);
      });
      node.addEventListener("mousemove", () => {
        const index = dock.items.findIndex((entry) => entry.node === node);
        if (index >= 0 && index !== dock.activeIndex) setCommandDockActive(index);
      });
      list.appendChild(node);
      dock.items.push({ ...item, node, group: group.key, section });
    });
    section.appendChild(list);
    const more = el("div", { className: "cmd-more reg-kick" });
    more.hidden = true;
    section.appendChild(more);
    if (!group.items.length && group.empty) {
      section.appendChild(el("div", { className: "cmd-empty", textContent: group.empty }));
    }
    dock.sections.push({ key: group.key, section, limit: group.limit || 0, more });
    host.appendChild(section);
  });
  filterCommandDock();
}

function getRailForKind(kind) {
  if (kind === "late") return "hatch";
  if (kind === "nostatus") return "dashed";
  if (kind === "done") return "muted";
  return "accent";
}

// Runs now when a deliverable is selected; otherwise parks the action and
// waits for the next selection.
function runCommandDockItem(item) {
  const dock = industryCommandDock;
  if (!item) return;
  if (item.kind === "project" || item.kind === "deliverable") {
    dock.input.value = "";
    item.run?.();
    return;
  }
  if (!dock.deliverable) {
    dock.pendingKey = item.key;
    dock.pendingLabel = item.label;
    dock.input.value = "";
    filterCommandDock();
    updateCommandDockContext();
    toast(`${item.label} — now select a deliverable.`);
    return;
  }
  const target = { deliverable: dock.deliverable, project: dock.project };
  dock.pendingKey = null;
  dock.pendingLabel = "";
  dock.input.value = "";
  filterCommandDock();
  updateCommandDockContext();
  const result = item.run?.(target);
  if (result && typeof result.catch === "function") {
    result.catch((error) => {
      console.warn("Command failed:", error);
      toast(error?.message || "Command failed.");
    });
  }
}

function getVisibleCommandDockItems() {
  return industryCommandDock.items.filter((item) => !item.node.hidden);
}

function setCommandDockActive(index) {
  const dock = industryCommandDock;
  dock.items.forEach((item) => item.node.classList.remove("is-active"));
  dock.activeIndex = -1;
  const item = dock.items[index];
  if (!item || item.node.hidden) return;
  item.node.classList.add("is-active");
  dock.activeIndex = index;
  if (dock.expanded) item.node.scrollIntoView?.({ block: "nearest" });
}

function moveCommandDockActive(step) {
  const dock = industryCommandDock;
  const visible = getVisibleCommandDockItems();
  if (!visible.length) return;
  const currentVisibleIndex = visible.findIndex((item) => item === dock.items[dock.activeIndex]);
  let nextVisible = currentVisibleIndex + step;
  if (nextVisible < 0) nextVisible = visible.length - 1;
  if (nextVisible >= visible.length) nextVisible = 0;
  setCommandDockActive(dock.items.indexOf(visible[nextVisible]));
}

function filterCommandDock() {
  const dock = industryCommandDock;
  const query = dock.input.value.trim().toLowerCase();
  const tokens = query.split(/\s+/).filter(Boolean);
  const matches = (item) => {
    if (!tokens.length) return true;
    const haystack = `${item.search || ""} ${item.label}`.toLowerCase();
    return tokens.every((token) => haystack.includes(token));
  };
  dock.items.forEach((item) => {
    item.node.hidden = !matches(item);
  });
  // Long target lists show only the first few until the query narrows them.
  (dock.sections || []).forEach(({ section, limit, more }) => {
    if (!limit) return;
    const sectionItems = dock.items.filter((item) => item.section === section && !item.node.hidden);
    sectionItems.forEach((item, index) => {
      if (index >= limit) item.node.hidden = true;
    });
    const overflow = sectionItems.length - limit;
    more.hidden = overflow <= 0;
    if (overflow > 0) {
      more.textContent = `+${overflow} more · keep typing to narrow`;
    }
  });
  (dock.sections || []).forEach(({ section }) => {
    const anyVisible = dock.items.some((item) => item.section === section && !item.node.hidden);
    const hasEmptyNote = !!section.querySelector(".cmd-empty");
    section.hidden = !anyVisible && !(hasEmptyNote && !query);
  });
  const visible = getVisibleCommandDockItems();
  const current = dock.items[dock.activeIndex];
  if (query) {
    if (!current || current.node.hidden) {
      setCommandDockActive(visible.length ? dock.items.indexOf(visible[0]) : -1);
    }
  } else if (current && current.node.hidden) {
    setCommandDockActive(-1);
  }
}

function focusCommandDockGroup(focus) {
  const dock = industryCommandDock;
  let index = -1;
  if (focus === "status") {
    index = dock.items.findIndex((item) => item.kind === "status" && item.checked);
    if (index < 0) index = dock.items.findIndex((item) => item.kind === "status");
  } else if (focus === "tools") {
    index = dock.items.findIndex((item) => item.group === "tools-frequent");
  }
  if (index >= 0) setCommandDockActive(index);
}

function handleCommandDockKeydown(event) {
  const dock = industryCommandDock;
  if (!dock) return;
  if (event.key === "ArrowDown") {
    event.preventDefault();
    if (!dock.expanded) setCommandDockExpanded(true);
    moveCommandDockActive(1);
    return;
  }
  if (event.key === "ArrowUp") {
    event.preventDefault();
    if (!dock.expanded) setCommandDockExpanded(true);
    moveCommandDockActive(-1);
    return;
  }
  if (event.key === "Enter") {
    event.preventDefault();
    if (event.ctrlKey || event.metaKey) {
      const quick = dock.items.find((item) => item.shortcut === "Enter");
      if (quick) runCommandDockItem(quick);
      return;
    }
    const active = dock.items[dock.activeIndex];
    if (active && !active.node.hidden) {
      runCommandDockItem(active);
      return;
    }
    const visible = getVisibleCommandDockItems();
    if (dock.input.value.trim() && visible.length === 1) runCommandDockItem(visible[0]);
    return;
  }
  if (event.key === "Escape") {
    event.preventDefault();
    if (dock.input.value) {
      dock.input.value = "";
      filterCommandDock();
    } else if (dock.pendingKey) {
      dock.pendingKey = null;
      dock.pendingLabel = "";
      updateCommandDockContext();
    } else if (dock.deliverable || dock.project) {
      clearCommandDockSelection();
    } else if (dock.expanded) {
      setCommandDockExpanded(false);
    } else {
      dock.input.blur();
    }
    return;
  }
  if ((event.ctrlKey || event.metaKey) && /^[1-9]$/.test(event.key)) {
    const target = dock.items.find((item) => item.shortcut === event.key);
    if (target) {
      event.preventDefault();
      runCommandDockItem(target);
    }
  }
}

// Selection ---------------------------------------------------------------

function selectDeliverableForCommands(deliverable, project, { expand = false, focus = "" } = {}) {
  const dock = ensureCommandDock();
  const same = dock.deliverable === deliverable;
  dock.deliverable = deliverable || null;
  dock.project = project || null;
  renderCommandDockItems();
  updateCommandDockContext();
  syncCommandDockSelection();

  if (dock.pendingKey && dock.deliverable) {
    const pending = dock.items.find((item) => item.key === dock.pendingKey);
    if (pending) {
      runCommandDockItem(pending);
      return;
    }
    dock.pendingKey = null;
    dock.pendingLabel = "";
    updateCommandDockContext();
  }

  if (expand) setCommandDockExpanded(true);
  if (focus) focusCommandDockGroup(focus);
  if (expand || focus) dock.input.focus({ preventScroll: true });
  else if (!same) dock.input.focus({ preventScroll: true });
}

// Picking a project narrows the line to that project's deliverables. A
// project with a single deliverable selects it straight away.
function selectProjectForCommands(project, { expand = true } = {}) {
  const dock = ensureCommandDock();
  const deliverables = getCommandDockProjectDeliverables(project);
  if (deliverables.length === 1) {
    selectDeliverableForCommands(deliverables[0], project, { expand });
    return;
  }
  dock.project = project || null;
  dock.deliverable = null;
  dock.input.value = "";
  renderCommandDockItems();
  updateCommandDockContext();
  syncCommandDockSelection();
  if (expand) setCommandDockExpanded(true);
  const first = dock.items.findIndex((item) => item.kind === "deliverable");
  if (first >= 0) setCommandDockActive(first);
  dock.input.focus({ preventScroll: true });
}

// Steps back: deliverable → project → nothing.
function clearCommandDockSelection() {
  const dock = ensureCommandDock();
  if (dock.deliverable) {
    dock.deliverable = null;
  } else {
    dock.project = null;
  }
  dock.input.value = "";
  renderCommandDockItems();
  updateCommandDockContext();
  syncCommandDockSelection();
}

// Re-applies the highlight after every render and drops selections whose
// records no longer exist.
function syncCommandDockSelection() {
  const dock = industryCommandDock;
  if (!dock) return;
  document
    .querySelectorAll(".reg-row.is-selected, .reg-line.is-selected, .plate-card.is-selected")
    .forEach((node) => node.classList.remove("is-selected"));
  if (dock.project && !(Array.isArray(db) && db.includes(dock.project))) {
    dock.project = null;
    dock.deliverable = null;
    renderCommandDockItems();
    updateCommandDockContext();
    return;
  }
  if (!dock.deliverable) return;
  const stillExists = getCommandDockProjectDeliverables(dock.project).includes(dock.deliverable);
  if (!stillExists) {
    dock.deliverable = null;
    renderCommandDockItems();
    updateCommandDockContext();
    return;
  }
  const id = String(dock.deliverable.id || "").trim();
  if (!id) return;
  document
    .querySelectorAll(`[data-deliverable-id="${CSS.escape(id)}"]`)
    .forEach((node) => {
      if (node.matches(".reg-row, .reg-line, .plate-card")) node.classList.add("is-selected");
    });
}

// Ctrl+K focuses the command line from anywhere on the Projects tab.
function bindCommandDockShortcut() {
  document.addEventListener("keydown", (event) => {
    if (!(event.ctrlKey || event.metaKey) || event.key.toLowerCase() !== "k") return;
    if (document.body.dataset.activeTab !== "projects") return;
    if (document.querySelector("dialog[open]")) return;
    const target = event.target;
    if (
      target &&
      target !== industryCommandDock?.input &&
      (target.matches?.("input, textarea, select, [contenteditable='true']") ||
        target.isContentEditable)
    ) {
      return;
    }
    event.preventDefault();
    const dock = ensureCommandDock();
    if (document.activeElement === dock.input && dock.expanded) {
      setCommandDockExpanded(false);
      dock.input.blur();
      return;
    }
    setCommandDockExpanded(true);
    dock.input.focus({ preventScroll: true });
    dock.input.select();
  });
}

// ---------------------------------------------------------------------------
// Chrome: footer and wiring
// ---------------------------------------------------------------------------

function updateProjectsIndustryChrome() {
  const footer = document.getElementById("registerFooter");
  if (footer) footer.hidden = projectsViewMode !== "list";
  // Projects and deliverables change with every render, so the dock's target
  // list is rebuilt here (the query and expansion state are kept).
  if (industryCommandDock) renderCommandDockItems();
  // The board renders after this call inside render(); re-apply the highlight
  // once it has.
  requestAnimationFrame(syncCommandDockSelection);
}

document.addEventListener("DOMContentLoaded", () => {
  ensureCommandDock();
  bindCommandDockShortcut();
});
