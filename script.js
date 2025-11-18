// ===================== CONFIGURATION & CONSTANTS =====================
const STATUS_CANON = ["Pending Review", "Complete", "Waiting"];
const LABEL_TO_KEY = { 'Pending Review': 'pendingReview', 'Complete': 'complete', 'Waiting': 'waiting' };
const KEY_TO_LABEL = { pendingReview: 'Pending Review', complete: 'Complete', waiting: 'Waiting' };

// Release meta data for bundles (Static Data)
const RELEASE_META = {
    "assets": [
        {
            "commands": {
                "EMBEDFROMXREFS": "Embeds raster images from XREFs into drawing (OLE/PowerPoint).",
                "EMBEDFROMPDFS": "Embeds PDF underlays into drawing (PNG -> OLE).",
                "CLEANTBLK": "Explodes blocks, keeps title block, detaches XREFs.",
                "CLEANCAD": "Embeds XREFs and performs cleanup."
            },
            "video_url": "https://loom.com/clean-cad-commands-demo",
            "filename": "ElectricalCommands.CleanCADCommands-v0.0.0.zip"
        },
        {
            "commands": {
                "WIPEOUTOBJECTS": "Creates wipeout objects behind text/tables.",
                "CREATEVIEWPORTFROMREGION": "Creates viewport in paperspace from modelspace region.",
                "CREATEPADDEDOUTLINEAROUNDBLOCKS": "Creates convex hull boundary around blocks.",
                "BLOCKBOUNDARYGRID": "Creates grid-based union boundary around blocks."
            },
            "video_url": "https://loom.com/general-commands-demo",
            "filename": "ElectricalCommands.GeneralCommands-v0.0.0.zip"
        },
        {
            "commands": {},
            "video_url": null,
            "filename": "ElectricalCommands.GetAttributesCommands-v0.0.0.zip"
        },
        {
            "commands": {
                "P30": "Plot layout to PDF (30x42).",
                "P24": "Plot layout to PDF (24x36).",
                "P22": "Plot layout to PDF (22x34)."
            },
            "video_url": "https://loom.com/plot-commands-demo",
            "filename": "ElectricalCommands.PlotCommands-v0.0.0.zip"
        },
        {
            "commands": {
                "INSERTPNGIMAGES": "Inserts PNGs in grid layout.",
                "INSERTPDFSHEETS": "Inserts PDF pages as underlays.",
                "GETSUMFROMTEXTEXPORT": "Sums areas/types and exports to JSON.",
                "GETSUMFROMTEXT": "Sums numerical values from text.",
                "CALCULATEAREAS": "Calculates polyline areas and labels them."
            },
            "video_url": "https://loom.com/t24-commands-demo",
            "filename": "ElectricalCommands.T24Commands-v0.0.0.zip"
        },
        {
            "commands": {
                "REPLACETEXTCONTENT": "Batch replace text content.",
                "INCREMENTTEXTCONTENT": "Increment text with prefixes/sorting.",
                "ADDVALUETOTEXT": "Add integer values to existing text numbers."
            },
            "video_url": "https://loom.com/text-commands-demo",
            "filename": "ElectricalCommands.TextCommands-v0.0.0.zip"
        }
    ]
};

// ===================== STATE MANAGEMENT =====================
/** @typedef {{ label:string, url:string, raw?:string }} Link */
/** @typedef {{ text:string, done?:boolean, links:Link[] }} Task */
/** @typedef {{ id:string, name:string, nick?:string, notes?:string, due?:string, path?:string, tasks:Task[], refs:Link[], statuses?:string[], statusTags?:string[], status?:string }} Project */

let db = [];
let notesDb = {}; // Simplified: flat object
let noteTabs = [];
let editIndex = -1;
let currentSort = { key: 'due', dir: 'desc' };
let statusFilter = 'all';
let dueFilter = 'all';

let userSettings = {
    userName: '',
    discipline: 'Electrical',
    apiKey: ''
};

// Notes State
let activeNoteTab = null;

// ===================== UTILITIES & HELPERS =====================

/** Create DOM Element helper */
function el(tag, props = {}, children = []) {
    const n = document.createElement(tag);
    Object.entries(props).forEach(([k, v]) => {
        if (k.startsWith('aria-') || k.startsWith('data-')) {
            if (v != null) n.setAttribute(k, String(v));
        } else {
            n[k] = v;
        }
    });
    children.forEach(c => {
        if (typeof c === 'string') n.appendChild(document.createTextNode(c));
        else n.appendChild(c);
    });
    return n;
}

/** Get Input Value helper */
function val(id) {
    const el = document.getElementById(id);
    return el ? el.value.trim() : '';
}

/** Close Dialog helper */
function closeDlg(id) {
    document.getElementById(id).close();
}

/** Debounce function for search/autosave */
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
    // ISO format
    let m = s.match(/^\d{4}-\d{2}-\d{2}$/);
    if (m) return new Date(s + 'T12:00:00');
    // US Format
    s = s.replace(/[.]/g, '/').replace(/\s+/g, '');
    const parts = s.split('/');
    if (parts.length === 3) {
        let [mm, dd, yy] = parts;
        if (yy.length === 2) yy = '20' + yy;
        const iso = `${yy}-${mm.padStart(2, '0')}-${dd.padStart(2, '0')}T12:00:00`;
        const d = new Date(iso);
        if (!isNaN(d)) return d;
    }
    const d2 = new Date(s);
    return isNaN(d2) ? null : d2;
}

function dueState(dueStr) {
    const d = parseDueStr(dueStr);
    if (!d) return 'ok';
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const dayOfWeek = today.getDay(); // 0 (Sun) to 6 (Sat)
    // Calculate start of current week (Sunday)
    const startOfWeek = new Date(today);
    startOfWeek.setDate(today.getDate() - dayOfWeek);
    // Calculate end of current week (Saturday)
    const endOfWeek = new Date(startOfWeek);
    endOfWeek.setDate(startOfWeek.getDate() + 6);
    endOfWeek.setHours(23, 59, 59, 999);

    if (d < today) return 'overdue';
    if (d >= startOfWeek && d <= endOfWeek) return 'dueSoon';
    return 'ok';
}

function humanDate(s) {
    const d = parseDueStr(s);
    if (!d) return '';
    return d.toLocaleDateString(undefined, { year: '2-digit', month: '2-digit', day: '2-digit' });
}

function basename(path) {
    try {
        if (!path) return '';
        const norm = path.replace(/\\/g, '/');
        const idx = norm.lastIndexOf('/');
        return idx >= 0 ? norm.slice(idx + 1) : norm;
    } catch { return path; }
}

function toFileURL(raw) {
    if (!raw) return '';
    let s = raw.trim();
    if (/^https?:\/\//i.test(s)) return s;
    if (/^\\\\/.test(s)) return 'file:' + s.replace(/^\\\\/, '/////').replace(/\\/g, '/');
    if (/^[A-Za-z]:\\/.test(s)) return 'file:///' + s.replace(/\\/g, '/');
    return s;
}

function normalizeLink(input) {
    const raw = (input || '').trim();
    const url = toFileURL(raw);
    const label = basename(raw) || raw || 'link';
    return { label, url, raw };
}

function convertPath(raw) {
    // ACies specific mapping
    if (raw.startsWith('\\\\acies.lan\\cachedrive\\projects2\\')) return raw.replace('\\\\acies.lan\\cachedrive\\projects2\\', 'M:\\');
    if (raw.startsWith('\\\\acies.lan\\cachedrive\\projects\\')) return raw.replace('\\\\acies.lan\\cachedrive\\projects\\', 'P:\\');
    return raw;
}

/** Styled Toast Notification */
function toast(msg, duration = 2500) {
    const existing = document.querySelector('.toast-notification');
    if (existing) existing.remove();

    const t = el('div', {
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
        `
    });

    // Add a tiny animation keyframe dynamically if needed, or rely on simple transition
    document.body.append(t);
    setTimeout(() => {
        t.style.opacity = '0';
        t.style.transform = 'translateX(-50%) translateY(10px)';
        t.style.transition = 'all 0.3s ease';
        setTimeout(() => t.remove(), 300);
    }, duration);
}

function resetToolStatus(toolId) {
    const card = document.getElementById(toolId);
    if (card) {
        card.classList.remove('running');
        const statusEl = card.querySelector('.tool-card-status');
        if (statusEl) {
            statusEl.textContent = '';
            statusEl.classList.remove('error');
        }
    }
}

// ===================== SERVER I/O (PYWEBVIEW) =====================

async function load() {
    try {
        const arr = await window.pywebview.api.get_tasks();
        migrateStatuses(arr);
        return arr;
    } catch (e) {
        console.warn('Backend load failed:', e);
        toast('⚠️ Offline Mode: Could not load data.');
        return [];
    }
}

async function save() {
    try {
        const response = await window.pywebview.api.save_tasks(db);
        if (response.status !== 'success') throw new Error(response.message);
    } catch (e) {
        console.warn('Backend save failed:', e);
        toast('⚠️ Failed to save data.');
    }
}

async function loadNotes() {
    try {
        const data = await window.pywebview.api.get_notes() || {};
        noteTabs = Array.isArray(data.tabs) && data.tabs.length > 0 ? data.tabs : ['General'];

        // Migration logic: Flatten Keyed/General into single object
        notesDb = {};
        noteTabs.forEach(tab => {
            const keyedContent = data.keyed && data.keyed[tab] ? data.keyed[tab].trim() : '';
            const generalContent = data.general && data.general[tab] ? data.general[tab].trim() : '';

            if (keyedContent && generalContent) {
                notesDb[tab] = `${keyedContent}\n\n--- General Notes ---\n\n${generalContent}`;
            } else {
                notesDb[tab] = keyedContent || generalContent;
            }
        });

        activeNoteTab = noteTabs[0];
        return notesDb;
    } catch (e) {
        console.warn('Backend notes load failed:', e);
        noteTabs = ['General'];
        notesDb = {};
        activeNoteTab = noteTabs[0];
        return {};
    }
}

async function saveNotes() {
    try {
        // Map flat notesDb to 'general' field for backend compatibility
        const dataToSave = {
            tabs: noteTabs,
            keyed: {}, // Deprecated field sent as empty
            general: notesDb
        };
        const response = await window.pywebview.api.save_notes(dataToSave);
        if (response.status !== 'success') throw new Error(response.message);
    } catch (e) {
        console.warn('Backend notes save failed:', e);
        toast('⚠️ Failed to save notes.');
    }
}

async function loadUserSettings() {
    try {
        const storedSettings = await window.pywebview.api.get_user_settings();
        if (storedSettings) userSettings = storedSettings;
    } catch (e) {
        console.error("Failed to load settings:", e);
    }
}

async function saveUserSettings() {
    try {
        await window.pywebview.api.save_user_settings(userSettings);
    } catch (e) {
        console.error("Failed to save settings:", e);
        toast('⚠️ Could not save settings.');
    }
}
const debouncedSaveUserSettings = debounce(saveUserSettings, 500);

// ===================== DATA MIGRATION & LOGIC =====================

function canonStatus(s) {
    if (!s) return null;
    const t = String(s).trim().toLowerCase();
    if (['pending review', 'pending-review', 'review', 'pr'].includes(t)) return 'Pending Review';
    if (['complete', 'completed', 'done'].includes(t)) return 'Complete';
    if (['waiting', 'wait'].includes(t)) return 'Waiting';
    return null;
}

function hasStatus(p, s) {
    return Array.isArray(p.statuses) && p.statuses.includes(s);
}

function toggleStatus(p, label) {
    if (!Array.isArray(p.statuses)) p.statuses = [];
    const isOn = p.statuses.includes(label);
    const key = LABEL_TO_KEY[label];
    if (key) setTag(p, key, !isOn);
}

function syncStatusArrays(p) {
    if (!Array.isArray(p.statuses)) p.statuses = [];
    const fromTags = Array.isArray(p.statusTags) ? p.statusTags : [];
    for (const k of fromTags) {
        const L = KEY_TO_LABEL[k];
        if (L && !p.statuses.includes(L)) p.statuses.push(L);
    }
    p.statuses = [...new Set(p.statuses.filter(s => STATUS_CANON.includes(s)))];
    p.statusTags = p.statuses.map(s => LABEL_TO_KEY[s]).filter(Boolean);

    // Legacy status string update
    if (p.statuses.includes('Complete')) p.status = 'Complete';
    else if (p.statuses.includes('Waiting')) p.status = 'Waiting';
    else if (p.statuses.includes('Pending Review')) p.status = 'Pending';
    else p.status = p.status || '';
}

function migrateStatuses(arr) {
    for (const p of arr) {
        if (!Array.isArray(p.statuses)) p.statuses = [];
        if (p.status) {
            String(p.status).split(/[,/|;]+/).map(s => s.trim()).filter(Boolean).forEach(piece => {
                const c = canonStatus(piece);
                if (c && !p.statuses.includes(c)) p.statuses.push(c);
            });
        }
        p.statuses = p.statuses.filter(s => STATUS_CANON.includes(s));
        syncStatusArrays(p);
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

    if (p.statuses.includes('Complete')) p.status = 'Complete';
    else if (p.statuses.includes('Waiting')) p.status = 'Waiting';
    else if (p.statuses.includes('Pending Review')) p.status = 'Pending';
    else p.status = '';
}

function getStatusTags(p) {
    let tags = Array.isArray(p.statusTags) ? [...p.statusTags] : [];
    const s = (p.status || '').toLowerCase();
    if (s) {
        if (s.includes('complete') && !tags.includes('complete')) tags.push('complete');
        if (s.includes('waiting') && !tags.includes('waiting')) tags.push('waiting');
        if ((s.includes('pending review') || s === 'pending') && !tags.includes('pendingReview')) tags.push('pendingReview');
    }
    return [...new Set(tags)];
}

// ===================== RENDER LOGIC =====================

function render() {
    const tbody = document.getElementById('tbody');
    const emptyState = document.getElementById('emptyState');
    tbody.innerHTML = '';

    const q = val('search').toLowerCase();

    // Filter items
    let items = db.filter(p => {
        if (q && !matches(q, p)) return false;

        // Due Filters
        if (dueFilter === 'overdue' && (dueState(p.due) !== 'overdue' || hasStatus(p, 'Complete'))) return false;
        if (dueFilter === 'soon' && dueState(p.due) !== 'dueSoon') return false;
        if (dueFilter === 'ok' && dueState(p.due) !== 'ok') return false;

        // Status Filters
        if (statusFilter === 'incomplete') {
            if (hasStatus(p, 'Complete')) return false;
        } else if (statusFilter !== 'all' && !hasStatus(p, statusFilter)) {
            return false;
        }
        return true;
    });

    // Sort items
    items.sort((a, b) => {
        const valA = a[currentSort.key];
        const valB = b[currentSort.key];
        let comparison = 0;
        if (currentSort.key === 'due') {
            const da = parseDueStr(valA), dbb = parseDueStr(valB);
            if (!da && !dbb) comparison = 0;
            else if (!da) comparison = 1;
            else if (!dbb) comparison = -1;
            else comparison = da - dbb;
        } else {
            comparison = String(valA || '').localeCompare(String(valB || ''), undefined, { numeric: true });
        }
        return comparison * (currentSort.dir === 'asc' ? 1 : -1);
    });

    updateSortHeaders();

    // Update KPI
    const incompleteDueSoon = db.filter(p => dueState(p.due) === 'dueSoon' && !hasStatus(p, 'Complete')).length;
    document.getElementById('kWeek').textContent = incompleteDueSoon;

    // Empty State Toggle
    emptyState.style.display = items.length ? 'none' : 'block';

    const rowTemplate = document.getElementById('project-row-template');

    items.forEach(p => {
        const tr = rowTemplate.content.cloneNode(true).querySelector('tr');
        const idx = db.indexOf(p);

        // ID (with styled badge)
        const idCell = tr.querySelector('.cell-id');
        const idBadge = idCell.querySelector('.id-badge') || idCell; // Fallback if template changed
        idBadge.textContent = p.id || '—';

        // Name & Path
        const nameCell = tr.querySelector('.cell-name');
        if (p.path) {
            const link = el('button', {
                className: 'path-link',
                textContent: p.name || '—',
                title: `Open: ${p.path}`
            });
            link.onclick = async () => {
                try {
                    await window.pywebview.api.open_path(convertPath(p.path));
                    toast('📂 Opening folder...');
                } catch (e) { toast('Failed to open path.'); }
            };
            nameCell.appendChild(link);
        } else {
            nameCell.textContent = p.name || '—';
        }
        if (p.nick) nameCell.append(el('small', { className: 'muted', textContent: ` (${p.nick})` }));

        // Due Date
        const dueCell = tr.querySelector('.cell-due');
        if (p.due) {
            const ds = dueState(p.due);
            const pillClass = ds === 'overdue' ? 'pill overdue' : ds === 'dueSoon' ? 'pill dueSoon' : 'pill ok';
            dueCell.appendChild(el('div', { className: pillClass, textContent: humanDate(p.due) }));
        } else {
            dueCell.textContent = '—';
        }

        // Status Toggles
        tr.querySelector('.cell-status').appendChild(renderStatusToggles(p));

        // Tasks (Expandable Logic)
        const taskCell = tr.querySelector('.cell-tasks');
        if (p.tasks && p.tasks.length) {
            const renderTasks = (expanded) => {
                taskCell.innerHTML = '';
                const tasksToShow = expanded ? p.tasks : p.tasks.slice(0, 2);

                tasksToShow.forEach(t => {
                    taskCell.appendChild(el('div', {
                        className: `task-chip ${t.done ? 'done' : ''}`,
                        textContent: t.text || 'Task'
                    }));
                });

                if (!expanded && p.tasks.length > 2) {
                    const moreBtn = el('button', {
                        className: 'btn-more-tasks',
                        textContent: `+${p.tasks.length - 2} more`,
                        onclick: (e) => {
                            e.stopPropagation();
                            renderTasks(true);
                        }
                    });
                    taskCell.appendChild(moreBtn);
                }

                // New: Minimize button if expanded
                if (expanded && p.tasks.length > 2) {
                    const lessBtn = el('button', {
                        className: 'btn-more-tasks',
                        textContent: `Show Less`,
                        style: 'margin-left: 8px',
                        onclick: (e) => {
                            e.stopPropagation();
                            renderTasks(false);
                        }
                    });
                    taskCell.appendChild(lessBtn);
                }
            };
            renderTasks(false);
        } else {
            taskCell.textContent = '—';
        }

        // Actions (Vertical)
        const actionsCell = tr.querySelector('.cell-actions');
        actionsCell.append(
            el('button', { className: 'btn', textContent: 'Edit', onclick: () => openEdit(idx) }),
            el('button', { className: 'btn', textContent: 'Duplicate', onclick: () => duplicate(idx) }),
            el('button', { className: 'btn btn-danger', textContent: 'Delete', onclick: () => removeProject(idx) }),
        );

        tbody.appendChild(tr);
    });
}

function renderStatusToggles(p) {
    const wrap = el('div', { className: 'status-group' });
    const mk = (cls, label) => {
        const b = el('button', {
            className: `st st-${cls}`,
            type: 'button',
            textContent: label.replace(' ', '\n'),
            title: label,
            'aria-pressed': String(hasStatus(p, label))
        });
        b.onclick = async (e) => {
            e.stopPropagation();
            toggleStatus(p, label);
            await save();
            render();
        };
        return b;
    };
    wrap.append(mk('pr', 'Pending Review'), mk('comp', 'Complete'), mk('wait', 'Waiting'));
    return wrap;
}

function matches(q, p) {
    if (!q) return true;
    const str = (val) => (val || '').toLowerCase();
    return str(p.id).includes(q) ||
        str(p.name).includes(q) ||
        str(p.nick).includes(q) ||
        str(p.notes).includes(q) ||
        (p.tasks || []).some(t => str(t.text).includes(q)) ||
        (p.statuses || []).some(s => str(s).includes(q));
}

function updateSortHeaders() {
    document.querySelectorAll('th[data-sort]').forEach(th => {
        th.classList.remove('sort-asc', 'sort-desc');
        if (th.dataset.sort === currentSort.key) {
            th.classList.add(`sort-${currentSort.dir}`);
        }
    });
}

// ===================== CRUD OPERATIONS =====================

function openEdit(i) {
    if (typeof i !== 'number' || i < 0 || !db[i]) return toast('Could not find row.');
    editIndex = i;
    const p = db[i];
    document.getElementById('dlgTitle').textContent = `Edit Project — ${p.id || 'Untitled'}`;
    document.getElementById('btnSaveProject').textContent = "Save Changes";
    fillForm(p);
    document.getElementById('editDlg').showModal();
}

function openNew() {
    editIndex = -1;
    document.getElementById('dlgTitle').textContent = 'New Project';
    document.getElementById('btnSaveProject').textContent = "Create Project";
    fillForm({ tasks: [], refs: [], statuses: [] });
    document.getElementById('editDlg').showModal();
}

function removeProject(i) {
    if (!confirm('Delete this project?')) return;
    const id = db[i]?.id;
    db.splice(i, 1);
    save();
    render();
    toast(`Deleted project ${id || ''}`);
}

function duplicate(i) {
    const original = db[i];
    const newProjectData = {
        id: original.id,
        name: original.name,
        nick: original.nick,
        path: original.path,
        refs: JSON.parse(JSON.stringify(original.refs || [])),
        due: '',
        notes: '',
        tasks: [],
        statuses: []
    };
    editIndex = -1;
    document.getElementById('dlgTitle').textContent = 'Duplicate Project';
    document.getElementById('btnSaveProject').textContent = "Create Duplicate";
    fillForm(newProjectData);
    document.getElementById('editDlg').showModal();
}

function markOverdueAsComplete() {
    let count = 0;
    db.forEach(p => {
        if (dueState(p.due) === 'overdue') {
            setTag(p, 'complete', true);
            count++;
        }
    });
    if (count > 0) {
        save();
        render();
        toast(`Marked ${count} overdue projects as complete.`);
    } else {
        toast('No overdue projects found.');
    }
}

// ===================== FORM HANDLING =====================

function fillForm(p) {
    document.getElementById('f_id').value = p.id || '';
    document.getElementById('f_name').value = p.name || '';
    document.getElementById('f_nick').value = p.nick || '';
    document.getElementById('f_notes').value = p.notes || '';
    document.getElementById('f_due').value = p.due || '';
    document.getElementById('f_path').value = p.path || '';
    setupStatusPicker('f_statuses', p.statuses || []);

    document.getElementById('taskList').innerHTML = '';
    (p.tasks || []).map(t => typeof t === 'string' ? { text: t } : t).forEach(addTaskRowFrom);

    document.getElementById('refList').innerHTML = '';
    (p.refs || []).forEach(addRefRowFrom);
}

function readForm() {
    const out = {
        id: val('f_id'),
        name: val('f_name'),
        nick: val('f_nick'),
        notes: val('f_notes'),
        due: val('f_due'),
        path: val('f_path'),
        tasks: [],
        refs: [],
        statuses: readStatusPicker('f_statuses')
    };

    // Tasks
    document.querySelectorAll('#taskList .task-row').forEach(row => {
        const text = row.querySelector('.t-text').value.trim();
        if (!text) return;
        const done = row.querySelector('.t-done').checked;
        const links = [row.querySelector('.t-link').value.trim(), row.querySelector('.t-link2').value.trim()]
            .filter(Boolean).map(normalizeLink);
        out.tasks.push({ text, done, links });
    });

    // References
    document.querySelectorAll('#refList .ref-row').forEach(row => {
        const label = row.querySelector('.r-label').value.trim();
        const raw = row.querySelector('.r-url').value.trim();
        if (!raw) return;
        const link = normalizeLink(raw);
        if (label) link.label = label;
        out.refs.push(link);
    });

    syncStatusArrays(out);
    return out;
}

function onSaveProject() {
    const data = readForm();
    if (editIndex >= 0) {
        db[editIndex] = data;
        toast('Project updated successfully.');
    } else {
        db.push(data);
        toast('Project created.');
    }
    save();
    render();
    closeDlg('editDlg');
}

function setupStatusPicker(containerId, selected) {
    const elc = document.getElementById(containerId);
    const setPressed = () => elc.querySelectorAll('.st').forEach(b => b.setAttribute('aria-pressed', String(selected.includes(b.dataset.status))));
    if (!elc.__wired) {
        elc.addEventListener('click', e => {
            if (e.target.matches('.st')) {
                const s = e.target.dataset.status;
                const i = selected.indexOf(s);
                if (i >= 0) selected.splice(i, 1);
                else selected.push(s);
                setPressed();
            }
        });
        elc.__wired = true;
    }
    setPressed();
}

function readStatusPicker(containerId) {
    return Array.from(document.querySelectorAll(`#${containerId} .st[aria-pressed="true"]`)).map(b => b.dataset.status);
}

function addTaskRowFrom(t = {}) {
    const template = document.getElementById('task-row-template');
    const row = template.content.cloneNode(true).querySelector('.task-row');
    row.querySelector('.t-text').value = t.text || '';
    row.querySelector('.t-done').checked = !!t.done;
    row.querySelector('.t-link').value = t.links?.[0]?.raw || '';
    row.querySelector('.t-link2').value = t.links?.[1]?.raw || '';
    row.querySelector('.btn-remove').onclick = () => row.remove();
    document.getElementById('taskList').append(row);
}

function addRefRowFrom(L = {}) {
    const template = document.getElementById('ref-row-template');
    const row = template.content.cloneNode(true).querySelector('.ref-row');
    row.querySelector('.r-label').value = L.label || '';
    row.querySelector('.r-url').value = L.raw || L.url || '';
    row.querySelector('.btn-remove').onclick = () => row.remove();
    document.getElementById('refList').append(row);
}

// Global accessors for button onclicks in HTML
window.addTaskRow = () => addTaskRowFrom({});
window.addRefRow = () => addRefRowFrom({});
window.closeDlg = closeDlg;

// ===================== CSV & IMPORT/EXPORT =====================

function parseCSV(text) {
    // Robust CSV parser handling quotes and newlines
    const rows = [];
    let i = 0, field = '', row = [], inQ = false;
    function pushField() { row.push(field); field = ''; }
    function pushRow() { rows.push(row); row = []; }

    while (i < text.length) {
        const ch = text[i];
        if (inQ) {
            if (ch === '"') {
                if (text[i + 1] === '"') { field += '"'; i += 2; continue; }
                inQ = false; i++; continue;
            }
            field += ch; i++; continue;
        } else {
            if (ch === '"') { inQ = true; i++; continue; }
            if (ch === ',') { pushField(); i++; continue; }
            if (ch === '\n') { pushField(); pushRow(); i++; continue; }
            if (ch === '\r') { if (text[i + 1] === '\n') { i += 2; pushField(); pushRow(); } else { i++; pushField(); pushRow(); } continue; }
            field += ch; i++;
        }
    }
    if (field !== '' || row.length) { pushField(); pushRow(); }
    return rows;
}

function importRows(rows, hasHeader = true) {
    if (rows.length && !hasHeader) {
        const joined = rows[0].map(s => (s || '').toUpperCase()).join(' | ');
        if (joined.includes('PROJECT NAME') || joined.includes('DUE')) hasHeader = true;
    }
    if (hasHeader) rows = rows.slice(1);

    let added = 0;
    for (const r of rows) {
        if (!r.length) continue;
        const [id, name, nick, notes, due, statusCell, tasksStr, path, ...refs] = r;
        if (!(id || name || tasksStr || refs.some(Boolean))) continue;

        const p = {
            id: String(id || '').trim(),
            name: (name || '').trim(),
            nick: (nick || '').trim(),
            notes: (notes || '').trim(),
            due: (due || '').trim(),
            path: (path || '').trim(),
            tasks: [],
            refs: [],
            statuses: []
        };

        // Status
        const parts = String(statusCell || '').split(/[,/|;]+/).map(s => s.trim()).filter(Boolean);
        for (const s of parts) {
            const c = canonStatus(s);
            if (c && !p.statuses.includes(c)) p.statuses.push(c);
        }

        // Tasks
        const tparts = (tasksStr || '').replace(/\r/g, '\n').split(/\n|;|\u2022|\r/).map(s => s.trim()).filter(Boolean);
        for (const t of tparts) p.tasks.push({ text: t, done: false, links: [] });

        // Refs
        for (const cell of refs) {
            const s = (cell || '').trim();
            if (!s) continue;
            p.refs.push(normalizeLink(s));
        }

        db.push(p);
        added++;
    }
    save();
    render();
    toast(`Imported ${added} rows`);
}

function exportCSV() {
    const header = ['ID', 'PROJECT NAME', 'NICKNAME', 'NOTES', 'DUE DATE', 'STATUS', 'TASKS', 'PATH', 'REFERENCE1', 'REFERENCE2', 'REFERENCE3', 'REFERENCE4'];
    const lines = [header.join(',')];
    for (const p of db) {
        const tasks = (p.tasks || []).map(t => t.text).join(' | ');
        const refs = (p.refs || []).map(L => L.raw || L.url).slice(0, 4);
        const statusStr = (p.statuses || []).join(' | ');
        const row = [p.id, p.name, p.nick || '', (p.notes || '').replaceAll('\n', ' '), p.due || '', statusStr, tasks, p.path || '', ...refs];
        const csv = row.map(cell => {
            const s = String(cell ?? '');
            return /[",\n]/.test(s) ? '"' + s.replaceAll('"', '""') + '"' : s;
        }).join(',');
        lines.push(csv);
    }
    const blob = new Blob([lines.join('\n')], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = el('a', { href: url, download: 'projects.csv' });
    document.body.append(a);
    a.click();
    a.remove();
}

function exportJSON() {
    const blob = new Blob([JSON.stringify(db, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = el('a', { href: url, download: 'projects.json' });
    document.body.append(a);
    a.click();
    a.remove();
}

// ===================== NOTES SYSTEM =====================
const debouncedSaveNotes = debounce(saveNotes, 500);

function renderNoteTabs() {
    const container = document.getElementById('notesTabsContainer');
    container.innerHTML = '';

    // Render existing tabs
    noteTabs.forEach(tabName => {
        const btn = el('button', {
            className: `inner-tab-btn ${tabName === activeNoteTab ? 'active' : ''}`,
            textContent: tabName,
            onclick: () => {
                activeNoteTab = tabName;
                renderNoteTabs();
            }
        });

        // Delete icon
        const delIcon = el('span', {
            className: 'tab-delete-icon',
            textContent: '🗑️',
            title: 'Delete Page',
            onclick: (e) => {
                e.stopPropagation(); // Prevent switching to the tab we're deleting
                if (confirm(`Permanently delete page "${tabName}"?`)) {
                    const idx = noteTabs.indexOf(tabName);
                    if (idx > -1) {
                        noteTabs.splice(idx, 1);
                        delete notesDb[tabName];
                        // If we deleted the active tab, switch to another one
                        if (activeNoteTab === tabName) {
                            activeNoteTab = noteTabs.length > 0 ? noteTabs[Math.max(0, idx - 1)] : null;
                        }
                        saveNotes();
                        renderNoteTabs();
                    }
                }
            }
        });

        btn.appendChild(delIcon);
        container.appendChild(btn);
    });

    // Add Page Button (+)
    const addBtn = el('button', {
        className: 'add-tab-btn',
        textContent: '+',
        title: 'Add New Page',
        onclick: () => {
            const name = prompt('Enter name for new page:');
            if (name && name.trim()) {
                if (!noteTabs.includes(name.trim())) {
                    noteTabs.push(name.trim());
                    activeNoteTab = name.trim();
                    saveNotes();
                    renderNoteTabs();
                } else {
                    toast('Page name already exists.');
                }
            }
        }
    });
    container.appendChild(addBtn);

    updateActiveNoteTextarea();
}

function updateActiveNoteTextarea() {
    const textarea = document.getElementById('notesTextarea');
    if (!activeNoteTab) {
        textarea.value = '';
        textarea.placeholder = 'Create a page to begin.';
        textarea.disabled = true;
        return;
    }
    textarea.disabled = false;
    textarea.placeholder = `Enter notes for ${activeNoteTab}...`;
    textarea.value = notesDb[activeNoteTab] || '';
}

function handleNoteInput(e) {
    if (!activeNoteTab) return;
    notesDb[activeNoteTab] = e.target.value;
    debouncedSaveNotes();
}

function renderNoteSearchResults() {
    const query = val('search').toLowerCase();
    const resultsContainer = document.getElementById('notesSearchResults');
    resultsContainer.innerHTML = '';
    if (!query) return;

    const queryWords = query.split(' ').filter(w => w);
    if (!queryWords.length) return;

    const seenNotes = new Set();

    for (const plan in notesDb) {
        const content = notesDb[plan];
        if (!content) continue;

        const notes = content.split(/\n\s*\n/).filter(note => note.trim() !== '');
        notes.forEach((noteText) => {
            const lowerNoteText = noteText.toLowerCase();
            if (queryWords.every(word => lowerNoteText.includes(word))) {
                const noteKey = `${plan}:${noteText}`;
                if (seenNotes.has(noteKey)) return;
                seenNotes.add(noteKey);

                const location = `${plan}`;
                const item = el('div', { className: 'note-result-item' });
                const contentDiv = el('div', { className: 'note-result-content' });
                const actionsDiv = el('div', { className: 'note-result-actions' });

                contentDiv.append(
                    el('div', { className: 'location', textContent: location }),
                    el('div', { className: 'snippet', textContent: noteText })
                );

                const editBtn = el('button', { className: 'btn tiny', textContent: 'Edit' });
                editBtn.onclick = () => {
                    activeNoteTab = plan;
                    renderNoteTabs();
                    setTimeout(() => {
                        const area = document.getElementById('notesTextarea');
                        if (area) {
                            area.focus();
                            const pos = area.value.indexOf(noteText);
                            if (pos !== -1) area.setSelectionRange(pos, pos);
                        }
                    }, 100);
                };

                const copyBtn = el('button', { className: 'btn tiny', textContent: 'Copy' });
                copyBtn.onclick = () => {
                    navigator.clipboard.writeText(noteText).then(() => toast('Copied!'));
                };

                actionsDiv.append(editBtn, copyBtn);
                item.append(contentDiv, actionsDiv);
                resultsContainer.append(item);
            }
        });
    }
}

// ===================== BUNDLE / PLUGIN MANAGER =====================

function openDetailsModal(asset) {
    const dlg = document.getElementById('commandDetailsDlg');
    if (!dlg || !asset) return;

    document.getElementById('detailsTitle').textContent = asset.filename.replace('ElectricalCommands.', '').replace(/-v[\d.]+\.zip$/, '');

    const videoEl = document.getElementById('detailsVideo');
    videoEl.innerHTML = '';
    if (asset.video_url && asset.video_url.includes('loom.com')) {
        const videoId = asset.video_url.split('/').pop();
        videoEl.append(el('iframe', { src: `https://www.loom.com/embed/${videoId}`, allowfullscreen: true }));
    } else {
        videoEl.innerHTML = '<p class="muted" style="padding:1rem;text-align:center">No video available.</p>';
    }

    const commandsEl = document.getElementById('detailsCommands');
    commandsEl.innerHTML = '';
    if (asset.commands && Object.keys(asset.commands).length > 0) {
        const list = el('ul', {}, Object.entries(asset.commands).map(([cmd, desc]) =>
            el('li', {}, [el('strong', { textContent: cmd }), `: ${desc}`])
        ));
        commandsEl.append(el('div', { className: 'bundle-commands' }, [el('h4', { textContent: 'Commands' }), list]));
    }
    dlg.showModal();
}

async function loadAndRenderBundles() {
    const container = document.getElementById('commands-container');
    if (!container) return;
    container.innerHTML = '<div class="spinner">Loading...</div>';

    try {
        const response = await window.pywebview.api.get_bundle_statuses();
        if (response.status !== 'success') throw new Error(response.message);

        container.innerHTML = '';
        if (response.data.length === 0) {
            container.textContent = 'No command bundles found.';
            return;
        }

        response.data.forEach(bundle => {
            const asset = RELEASE_META.assets.find(a => a.filename.includes(bundle.name));
            const card = el('div', { className: 'release-card' });

            // Status determination
            let statusClass, statusTitle, btnText, btnClass;
            if (bundle.state === 'installed') {
                statusClass = 'installed'; statusTitle = `Installed (v${bundle.local_version})`;
                btnText = 'Uninstall'; btnClass = 'btn-danger';
            } else if (bundle.state === 'update_available') {
                statusClass = 'update-available'; statusTitle = `Update Available (v${bundle.remote_version})`;
                btnText = 'Update'; btnClass = 'btn-accent';
            } else {
                statusClass = 'not-installed'; statusTitle = 'Not Installed';
                btnText = 'Install'; btnClass = 'btn-primary';
            }

            // Header
            const header = el('div', { className: 'release-card-header' }, [
                el('div', { className: 'release-card-title' }, [
                    el('div', { className: `bundle-status ${statusClass}`, title: statusTitle }),
                    el('span', { textContent: bundle.name.replace('ElectricalCommands.', '') })
                ]),
                el('button', {
                    className: 'info-btn', textContent: '?', title: 'Details',
                    onclick: () => asset ? openDetailsModal(asset) : toast('No details available')
                })
            ]);

            // Body (Tags)
            const body = el('div', { className: 'release-card-body' });
            const tags = el('div', { className: 'command-tags' });
            if (asset && asset.commands) {
                Object.keys(asset.commands).forEach(cmd => tags.append(el('span', { className: 'command-tag', textContent: cmd })));
            }
            body.append(tags);

            // Footer (Action)
            const footer = el('div', { className: 'release-card-footer' });
            const btn = el('button', { className: `btn ${btnClass}`, textContent: btnText });
            btn.dataset.bundleName = bundle.bundle_name;
            btn.dataset.actionType = btnText;
            if (bundle.state !== 'installed') btn.dataset.asset = JSON.stringify(bundle.asset);
            footer.append(btn);

            card.append(header, body, footer);
            container.append(card);
        });
    } catch (e) {
        container.innerHTML = `<div class="error-message">Error: ${e.message}</div>`;
    }
}

async function handleBundleAction(e) {
    const button = e.target.closest('[data-action-type]');
    if (!button) return;

    const actionType = button.dataset.actionType;
    button.disabled = true;
    button.textContent = 'Processing...';

    try {
        let response;
        if (actionType === 'Install' || actionType === 'Update') {
            const asset = JSON.parse(button.dataset.asset);
            response = await window.pywebview.api.install_single_bundle(asset);
        } else {
            response = await window.pywebview.api.uninstall_bundle(button.dataset.bundleName);
        }
        if (response.status !== 'success') throw new Error(response.message);
        toast(`${actionType} successful.`);
    } catch (err) {
        toast(`⚠️ ${err.message}`, 5000);
    } finally {
        await loadAndRenderBundles();
        await updateCleanDwgToolState();
    }
}

// ===================== TOOLS & SCRIPTS =====================

// Python Callbacks
window.updateToolStatus = function (toolId, message) {
    const card = document.getElementById(toolId);
    if (!card) return;
    const statusEl = card.querySelector('.tool-card-status');
    const abortBtn = document.getElementById('abortBtn');

    statusEl.classList.remove('error');

    // Abort button logic
    if (toolId === 'toolCleanDwgs' && abortBtn) {
        if (message && message !== 'DONE' && !message.startsWith('ERROR:')) {
            abortBtn.style.display = 'inline-block';
            abortBtn.disabled = false;
            abortBtn.onclick = async () => {
                await window.pywebview.api.abort_clean_dwgs();
                toast('Aborting...');
                abortBtn.disabled = true;
            };
        } else {
            abortBtn.style.display = 'none';
        }
    }

    if (message.startsWith('ERROR:')) {
        statusEl.textContent = message.substring(6).trim();
        statusEl.classList.add('error');
        card.classList.add('running');
        setTimeout(() => {
            card.classList.remove('running');
            if (abortBtn) abortBtn.style.display = 'none';
        }, 5000);
    } else if (message === 'DONE') {
        statusEl.textContent = 'Done!';
        setTimeout(() => {
            card.classList.remove('running');
            if (abortBtn) abortBtn.style.display = 'none';
        }, 2000);
    } else {
        statusEl.textContent = message;
    }
};

async function updateCleanDwgToolState() {
    const toolCard = document.getElementById('toolCleanDwgs');
    if (!toolCard) return;
    try {
        const response = await window.pywebview.api.check_bundle_installed('ElectricalCommands.CleanCADCommands.bundle');
        if (response.status === 'success' && response.is_installed) {
            toolCard.classList.remove('disabled');
            toolCard.setAttribute('tabindex', '0');
        } else {
            toolCard.classList.add('disabled');
            toolCard.setAttribute('tabindex', '-1');
        }
    } catch (e) {
        toolCard.classList.add('disabled');
    }
}

// ===================== INITIALIZATION & EVENTS =====================

function initTabbedInterfaces() {
    // Main Nav Logic
    const mainTabContainer = document.querySelector('.main-nav');
    const searchInput = document.getElementById('search');

    mainTabContainer.addEventListener('click', e => {
        if (!e.target.matches('.main-tab-btn')) return;
        const tab = e.target.dataset.tab;

        // UI Updates
        document.querySelectorAll('.main-tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === tab));
        document.querySelectorAll('.tab-panel').forEach(p => {
            p.hidden = p.id !== `${tab}-panel`;
            p.classList.toggle('active', p.id === `${tab}-panel`);
        });

        // Context-aware actions
        if (tab === 'plugins') loadAndRenderBundles();

        if (tab === 'notes') {
            searchInput.placeholder = 'Search notes...';
            searchInput.disabled = false;
        } else if (tab === 'plugins') {
            searchInput.placeholder = 'Search unavailable';
            searchInput.disabled = true;
        } else {
            searchInput.placeholder = 'Search projects...';
            searchInput.disabled = false;
        }
        searchInput.value = '';
    });
}

function initEventListeners() {
    // Global Search
    document.getElementById('search').addEventListener('input', debounce(() => {
        const activeTab = document.querySelector('.main-tab-btn.active')?.dataset.tab;
        if (activeTab === 'notes') renderNoteSearchResults();
        else render();
    }, 250));

    // Toolbar Actions
    document.getElementById('quickNew').onclick = openNew;
    document.getElementById('btnNew').onclick = openNew;
    document.getElementById('settingsBtn').onclick = () => document.getElementById('settingsDlg').showModal();

    // Drawer Logic
    const drawer = document.getElementById('drawer');
    const backdrop = document.getElementById('backdrop');
    const toggleDrawer = (open) => {
        if (open) { drawer.classList.add('open'); backdrop.hidden = false; }
        else { drawer.classList.remove('open'); backdrop.hidden = true; }
    };
    document.getElementById('menuBtn').onclick = () => toggleDrawer(true);
    document.getElementById('drawerClose').onclick = () => toggleDrawer(false);
    backdrop.onclick = () => toggleDrawer(false);

    // Filters
    document.getElementById('dueFilterGroup').addEventListener('click', e => {
        if (e.target.matches('.filter-chip')) {
            dueFilter = e.target.dataset.dueFilter;
            document.querySelectorAll('#dueFilterGroup .filter-chip').forEach(b => b.classList.remove('active'));
            e.target.classList.add('active');
            render();
        }
    });
    document.getElementById('statusFilterGroup').addEventListener('click', e => {
        if (e.target.matches('.filter-chip')) {
            statusFilter = e.target.dataset.filter;
            document.querySelectorAll('#statusFilterGroup .filter-chip').forEach(b => b.classList.remove('active'));
            e.target.classList.add('active');
            render();
        }
    });

    // Tools
    document.getElementById('toolPublishDwgs').addEventListener('click', async (e) => {
        if (e.currentTarget.classList.contains('running')) return;
        e.currentTarget.classList.add('running');
        window.updateToolStatus('toolPublishDwgs', 'Initializing...');
        await window.pywebview.api.run_publish_script();
    });

    document.getElementById('toolCleanDwgs').addEventListener('click', (e) => {
        if (e.currentTarget.classList.contains('disabled')) {
            document.getElementById('installPrereqDlg').showModal();
            return;
        }
        if (e.currentTarget.classList.contains('running')) return;
        document.getElementById('cleanDwgsDlg').showModal();
    });

    document.getElementById('btnProceedCleanDwgs').addEventListener('click', async () => {
        const titleblock = val('cleanDwgs_titleblockPath');
        if (!titleblock) return toast('Select a titleblock first.');

        const disciplines = Array.from(document.querySelectorAll('input[name="cleanDwgs_discipline"]:checked')).map(c => c.value);
        closeDlg('cleanDwgsDlg');

        const response = await window.pywebview.api.run_clean_dwgs_script({ titleblock, disciplines });
        if (response.status === 'success') {
            window.updateToolStatus('toolCleanDwgs', 'Processing...');
            document.getElementById('toolCleanDwgs').classList.add('running');
        } else if (response.status === 'prerequisite_failed') {
            toast('⚠️ ' + response.message);
        }
    });

    // Clean DWGs Browse Button
    document.getElementById('btnSelectTitleblock').addEventListener('click', async () => {
        const res = await window.pywebview.api.select_files({ allow_multiple: false, file_types: ['AutoCAD Drawings (*.dwg)'] });
        if (res.status === 'success' && res.paths.length) {
            document.getElementById('cleanDwgs_titleblockPath').value = res.paths[0];
        }
    });

    // AI / Email Logic
    const handleAI = () => {
        if (!userSettings.apiKey) { toast('Setup API Key in Settings first.'); return; }
        document.getElementById('emailArea').value = '';
        document.getElementById('aiSpinner').style.display = 'none';
        document.getElementById('emailDlg').showModal();
    };
    document.getElementById('btnEmail').onclick = handleAI;
    document.getElementById('aiBtn').onclick = handleAI;

    document.getElementById('btnProcessEmail').onclick = async () => {
        const txt = val('emailArea');
        if (!txt) return;
        document.getElementById('aiSpinner').style.display = 'flex';
        try {
            const res = await window.pywebview.api.process_email_with_ai(txt, userSettings.apiKey, userSettings.userName, userSettings.discipline);
            if (res.status === 'success') {
                closeDlg('emailDlg');
                openNew();
                fillForm(res.data);
            } else throw new Error(res.message);
        } catch (e) { toast('AI Error: ' + e.message); }
        document.getElementById('aiSpinner').style.display = 'none';
    };

    // Notes Logic (Input Handler Only - buttons handled in renderNoteTabs)
    document.getElementById('notesTextarea').addEventListener('input', handleNoteInput);

    // Settings Inputs
    document.getElementById('settings_userName').oninput = (e) => { userSettings.userName = e.target.value; debouncedSaveUserSettings(); };
    document.getElementById('settings_apiKey').oninput = (e) => { userSettings.apiKey = e.target.value; debouncedSaveUserSettings(); };

    // Export/Import
    document.getElementById('btnExportCsv').onclick = exportCSV;
    document.getElementById('btnExportJson').onclick = exportJSON;
    document.getElementById('btnImport').onclick = async () => {
        const res = await window.pywebview.api.import_and_process_csv();
        if (res.status === 'success' && res.data) {
            db.push(...res.data);
            migrateStatuses(db);
            save();
            render();
            toast(`Imported ${res.data.length} projects.`);
        }
    };

    // Paste CSV
    document.getElementById('btnPaste').onclick = () => document.getElementById('pasteDlg').showModal();
    document.getElementById('btnPasteImport').onclick = () => {
        const rows = parseCSV(val('pasteArea'));
        importRows(rows, document.getElementById('hasHeader').checked);
        closeDlg('pasteDlg');
    };

    // Deletion
    document.getElementById('btnDeleteAll').onclick = () => {
        document.getElementById('deleteConfirmInput').value = '';
        document.getElementById('btnDeleteConfirm').disabled = true;
        document.getElementById('deleteDlg').showModal();
    };
    document.getElementById('deleteConfirmInput').oninput = (e) => {
        document.getElementById('btnDeleteConfirm').disabled = e.target.value !== 'DELETE';
    };
    document.getElementById('btnDeleteConfirm').onclick = () => {
        db = []; save(); render(); closeDlg('deleteDlg');
    };

    // Bundle Actions
    document.getElementById('commands-container').addEventListener('click', handleBundleAction);
    document.getElementById('btnInstallPrereq').onclick = async () => {
        // Simpler install flow for prereq
        const spinner = document.getElementById('installSpinner');
        spinner.style.display = 'block';
        try {
            const all = await window.pywebview.api.get_bundle_statuses();
            const bundle = all.data.find(b => b.name.includes('CleanCADCommands'));
            if (bundle) {
                await window.pywebview.api.install_single_bundle(bundle.asset);
                toast('Installed!');
                closeDlg('installPrereqDlg');
                updateCleanDwgToolState();
                loadAndRenderBundles();
            }
        } catch (e) { toast('Installation failed.'); }
        spinner.style.display = 'none';
    };
}

async function init() {
    // Ensure python backend is ready
    if (!window.pywebview) await new Promise(r => window.addEventListener('pywebviewready', r));

    initEventListeners();
    initTabbedInterfaces();

    await loadUserSettings();
    db = await load();
    await loadNotes();

    renderNoteTabs();
    render();

    // Initial check for tool availability
    updateCleanDwgToolState();
}

// Start
init();