// ===================== CONFIGURATION & CONSTANTS =====================
const STATUS_CANON = ["Waiting", "Working", "Pending Review", "Complete", "Delivered"];
const STATUS_PRIORITY = ['Delivered', 'Complete', 'Pending Review', 'Working', 'Waiting'];
const LABEL_TO_KEY = { 'Waiting': 'waiting', 'Working': 'working', 'Pending Review': 'pendingReview', 'Complete': 'complete', 'Delivered': 'delivered' };
const KEY_TO_LABEL = { waiting: 'Waiting', working: 'Working', pendingReview: 'Pending Review', complete: 'Complete', delivered: 'Delivered' };

// Cache for loaded descriptions to prevent re-fetching
const DESCRIPTION_CACHE = {};

// ===================== UTILITIES & HELPERS =====================

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

function val(id) {
    const el = document.getElementById(id);
    return el ? el.value.trim() : '';
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
    if (m) return new Date(s + 'T12:00:00');
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
    const dayOfWeek = today.getDay();
    const startOfWeek = new Date(today);
    startOfWeek.setDate(today.getDate() - dayOfWeek);
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
    if (raw.startsWith('\\\\acies.lan\\cachedrive\\projects2\\')) return raw.replace('\\\\acies.lan\\cachedrive\\projects2\\', 'M:\\');
    if (raw.startsWith('\\\\acies.lan\\cachedrive\\projects\\')) return raw.replace('\\\\acies.lan\\cachedrive\\projects\\', 'P:\\');
    return raw;
}

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

    document.body.append(t);
    setTimeout(() => {
        t.style.opacity = '0';
        t.style.transform = 'translateX(-50%) translateY(10px)';
        t.style.transition = 'all 0.3s ease';
        setTimeout(() => t.remove(), 300);
    }, duration);
}

const updateStickyOffsets = () => {
    const header = document.querySelector('.app-header');
    const toolbar = document.querySelector('#projects-panel .panel-toolbar');
    if (header) document.documentElement.style.setProperty('--header-height', `${header.offsetHeight}px`);
    if (toolbar) document.documentElement.style.setProperty('--toolbar-height', `${toolbar.offsetHeight}px`);
};

const debouncedStickyOffsets = debounce(updateStickyOffsets, 150);
window.addEventListener('resize', debouncedStickyOffsets);

// ===================== SERVER I/O =====================

// State variables
let db = [];
let notesDb = {};
let noteTabs = [];
let editIndex = -1;
let currentSort = { key: 'due', dir: 'desc' };
let statusFilter = 'all';
let dueFilter = 'all';

let userSettings = {
    userName: '',
    discipline: 'Electrical',
    apiKey: '',
    autocadPath: ''
};
let activeNoteTab = null;

async function load() {
    try {
        const arr = await window.pywebview.api.get_tasks();
        migrateStatuses(arr);
        return arr;
    } catch (e) {
        console.warn('Backend load failed:', e);
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
        noteTabs = ['General'];
        notesDb = {};
        activeNoteTab = noteTabs[0];
        return {};
    }
}

async function saveNotes() {
    try {
        const dataToSave = { tabs: noteTabs, keyed: {}, general: notesDb };
        await window.pywebview.api.save_notes(dataToSave);
    } catch (e) { console.warn('Backend notes save failed:', e); }
}

async function loadUserSettings() {
    try {
        const storedSettings = await window.pywebview.api.get_user_settings();
        if (storedSettings) userSettings = storedSettings;
    } catch (e) { console.error("Failed to load settings:", e); }
}

async function populateSettingsModal() {
    document.getElementById('settings_userName').value = userSettings.userName || '';
    document.getElementById('settings_apiKey').value = userSettings.apiKey || '';
    document.getElementById('settings_autocadPath').value = userSettings.autocadPath || '';
    const discipline = userSettings.discipline || 'Electrical';
    document.querySelectorAll('input[name="settings_discipline_radio"]').forEach(radio => {
        radio.checked = radio.value === discipline;
    });

    // Populate AutoCAD versions
    const container = document.getElementById('autocad_versions_container');
    container.innerHTML = '<div class="spinner">Detecting versions...</div>';
    try {
        const response = await window.pywebview.api.get_installed_autocad_versions();
        if (response.status === 'success') {
            container.innerHTML = '';
            response.versions.forEach(version => {
                const radio = el('input', {
                    type: 'radio',
                    name: 'autocad_version_radio',
                    value: version.path,
                    checked: userSettings.autocadPath === version.path
                });
                radio.onchange = () => {
                    userSettings.autocadPath = radio.value;
                    debouncedSaveUserSettings();
                };
                const label = el('label', { className: 'radio-label' }, [radio, ` AutoCAD ${version.year}`]);
                container.appendChild(label);
            });
            if (response.versions.length === 0) {
                container.innerHTML = '<p class="tiny muted">No AutoCAD versions detected in default location.</p>';
            }
        } else {
            container.innerHTML = '<p class="tiny muted">Error detecting versions.</p>';
        }
    } catch (e) {
        container.innerHTML = '<p class="tiny muted">Error detecting versions.</p>';
    }
}

async function populateAutocadSelectModal() {
    document.getElementById('autocad_select_custom').value = userSettings.autocadPath || '';
    const container = document.getElementById('autocad_select_container');
    container.innerHTML = '<div class="spinner">Detecting versions...</div>';
    try {
        const response = await window.pywebview.api.get_installed_autocad_versions();
        if (response.status === 'success') {
            container.innerHTML = '';
            response.versions.forEach(version => {
                const radio = el('input', {
                    type: 'radio',
                    name: 'autocad_select_radio',
                    value: version.path,
                    checked: userSettings.autocadPath === version.path
                });
                const label = el('label', { className: 'radio-label' }, [radio, ` AutoCAD ${version.year}`]);
                container.appendChild(label);
            });
            if (response.versions.length === 0) {
                container.innerHTML = '<p class="tiny muted">No AutoCAD versions detected in default location.</p>';
            }
        } else {
            container.innerHTML = '<p class="tiny muted">Error detecting versions.</p>';
        }
    } catch (e) {
        container.innerHTML = '<p class="tiny muted">Error detecting versions.</p>';
    }
}

async function showAutocadSelectModal() {
    await populateAutocadSelectModal();
    document.getElementById('autocadSelectDlg').showModal();
}

async function saveUserSettings() {
    // Update autocadPath from UI
    const selectedRadio = document.querySelector('input[name="autocad_version_radio"]:checked');
    if (selectedRadio) {
        userSettings.autocadPath = selectedRadio.value;
    } else {
        userSettings.autocadPath = document.getElementById('settings_autocadPath').value.trim();
    }
    try { await window.pywebview.api.save_user_settings(userSettings); }
    catch (e) { toast('⚠️ Could not save settings.'); }
}
const debouncedSaveUserSettings = debounce(saveUserSettings, 500);

// ===================== DATA MIGRATION =====================
function canonStatus(s) {
    if (!s) return null;
    const t = String(s).trim().toLowerCase();
    if (['waiting', 'wait', 'blocked'].includes(t)) return 'Waiting';
    if (['working', 'work', 'in progress', 'in-progress', 'doing', 'active'].includes(t)) return 'Working';
    if (['pending review', 'pending-review', 'review', 'pr', 'pending'].includes(t)) return 'Pending Review';
    if (['complete', 'completed', 'done'].includes(t)) return 'Complete';
    if (['delivered', 'sent', 'shipped'].includes(t)) return 'Delivered';
    return null;
}
function hasStatus(p, s) { return Array.isArray(p.statuses) && p.statuses.includes(s); }
function isFinished(p) { return hasStatus(p, 'Complete') || hasStatus(p, 'Delivered'); }
function applyPrimaryStatus(p) {
    const primary = STATUS_PRIORITY.find(status => hasStatus(p, status));
    p.status = primary || '';
}
function toggleStatus(p, label) {
    if (!Array.isArray(p.statuses)) p.statuses = [];
    const key = LABEL_TO_KEY[label];
    if (key) setTag(p, key, !p.statuses.includes(label));
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
    applyPrimaryStatus(p);
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
    applyPrimaryStatus(p);
}
function getStatusTags(p) {
    let tags = Array.isArray(p.statusTags) ? [...p.statusTags] : [];
    const s = (p.status || '').toLowerCase();
    if (s) {
        if (s.includes('complete') && !tags.includes('complete')) tags.push('complete');
        if (s.includes('waiting') && !tags.includes('waiting')) tags.push('waiting');
        if ((s.includes('working') || s.includes('in progress')) && !tags.includes('working')) tags.push('working');
        if ((s.includes('pending review') || s === 'pending') && !tags.includes('pendingReview')) tags.push('pendingReview');
        if (s.includes('deliver') && !tags.includes('delivered')) tags.push('delivered');
    }
    return [...new Set(tags)];
}

// ===================== RENDER LOGIC =====================
function render() {
    const tbody = document.getElementById('tbody');
    const emptyState = document.getElementById('emptyState');
    tbody.innerHTML = '';

    const q = val('search').toLowerCase();

    let items = db.filter(p => {
        if (q && !matches(q, p)) return false;
        const d = parseDueStr(p.due);

        if (dueFilter === 'past') {
            if (!d) return false;
            const today = new Date(); today.setHours(0, 0, 0, 0);
            const dayOfWeek = today.getDay();
            const startOfWeek = new Date(today); startOfWeek.setDate(today.getDate() - dayOfWeek); startOfWeek.setHours(0, 0, 0, 0);
            if (d >= startOfWeek) return false;
        }
        if (dueFilter === 'soon') {
            if (!d) return false;
            const today = new Date(); today.setHours(0, 0, 0, 0);
            const dayOfWeek = today.getDay();
            const startOfWeek = new Date(today); startOfWeek.setDate(today.getDate() - dayOfWeek); startOfWeek.setHours(0, 0, 0, 0);
            const endOfWeek = new Date(startOfWeek); endOfWeek.setDate(startOfWeek.getDate() + 6); endOfWeek.setHours(23, 59, 59, 999);
            if (d < startOfWeek || d > endOfWeek) return false;
        }
        if (dueFilter === 'future') {
            if (!d) return false;
            const today = new Date(); today.setHours(0, 0, 0, 0);
            const dayOfWeek = today.getDay();
            const startOfWeek = new Date(today); startOfWeek.setDate(today.getDate() - dayOfWeek);
            const endOfWeek = new Date(startOfWeek); endOfWeek.setDate(startOfWeek.getDate() + 6); endOfWeek.setHours(23, 59, 59, 999);
            if (d <= endOfWeek) return false;
        }
        if (statusFilter === 'incomplete') {
            if (isFinished(p)) return false;
        } else if (statusFilter !== 'all' && !hasStatus(p, statusFilter)) {
            return false;
        }
        return true;
    });

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

    const total = db.length;
    const completed = db.filter(p => isFinished(p)).length;
    const today = new Date(); today.setHours(0, 0, 0, 0);
    const dayOfWeek = today.getDay();
    const startOfWeek = new Date(today); startOfWeek.setDate(today.getDate() - dayOfWeek);
    const endOfWeek = new Date(startOfWeek); endOfWeek.setDate(startOfWeek.getDate() + 6); endOfWeek.setHours(23, 59, 59, 999);
    const startOfLastWeek = new Date(startOfWeek); startOfLastWeek.setDate(startOfWeek.getDate() - 7);
    const endOfLastWeek = new Date(endOfWeek); endOfLastWeek.setDate(endOfWeek.getDate() - 7);
    const currentYear = today.getFullYear();
    const lastYear = currentYear - 1;

    let dueThisWeek = 0, dueLastWeek = 0, upcoming = 0, completedThisYear = 0, completedLastYear = 0;
    let minDate = null, maxDate = null;

    db.forEach(p => {
        const d = parseDueStr(p.due);
        if (d) {
            if (d >= startOfWeek && d <= endOfWeek) dueThisWeek++;
            if (d >= startOfLastWeek && d <= endOfLastWeek) dueLastWeek++;
            if (d > endOfWeek) upcoming++;
            if (!minDate || d < minDate) minDate = d;
            if (!maxDate || d > maxDate) maxDate = d;
            if (isFinished(p)) {
                if (d.getFullYear() === currentYear) completedThisYear++;
                if (d.getFullYear() === lastYear) completedLastYear++;
            }
        }
    });

    let avgPerWeek = 0, avgPerMonth = 0;
    if (minDate && maxDate && total > 0) {
        const diffTime = Math.abs(maxDate - minDate);
        const diffWeeks = Math.ceil(diffTime / (1000 * 60 * 60 * 24 * 7)) || 1;
        avgPerWeek = (total / diffWeeks).toFixed(1);
        const diffMonths = (diffTime / (1000 * 60 * 60 * 24 * 30.44)) || 1;
        avgPerMonth = (total / diffMonths).toFixed(1);
    }

    const setStat = (id, val) => { const el = document.getElementById(id); if (el) el.textContent = val; };
    setStat('statTotal', total);
    setStat('statCompleted', completed);
    setStat('statLastWeek', dueLastWeek);
    setStat('statThisWeek', dueThisWeek);
    setStat('statUpcoming', upcoming);
    setStat('statAvg', avgPerWeek);
    setStat('statAvgMonth', avgPerMonth);
    setStat('statCompletedLastYear', completedLastYear);
    setStat('statCompletedThisYear', completedThisYear);

    emptyState.style.display = items.length ? 'none' : 'block';
    const rowTemplate = document.getElementById('project-row-template');

    items.forEach(p => {
        const tr = rowTemplate.content.cloneNode(true).querySelector('tr');
        const idx = db.indexOf(p);

        const idCell = tr.querySelector('.cell-id');
        const idBadge = idCell.querySelector('.id-badge') || idCell;
        idBadge.textContent = p.id || '—';

        const nameCell = tr.querySelector('.cell-name');
        if (p.path) {
            const link = el('button', {
                className: 'path-link',
                textContent: p.name || '—',
                title: `Open: ${p.path}`
            });
            link.onclick = async () => {
                try { await window.pywebview.api.open_path(convertPath(p.path)); toast('📂 Opening folder...'); }
                catch (e) { toast('Failed to open path.'); }
            };
            nameCell.appendChild(link);
        } else {
            nameCell.textContent = p.name || '—';
        }
        if (p.nick) nameCell.append(el('small', { className: 'muted', textContent: ` (${p.nick})` }));

        const dueCell = tr.querySelector('.cell-due');
        if (p.due) {
            const ds = dueState(p.due);
            const pillClass = ds === 'overdue' ? 'pill overdue' : ds === 'dueSoon' ? 'pill dueSoon' : 'pill ok';
            dueCell.appendChild(el('div', { className: pillClass, textContent: humanDate(p.due) }));
        } else {
            dueCell.textContent = '—';
        }

        tr.querySelector('.cell-status').appendChild(renderStatusToggles(p));

        const taskCell = tr.querySelector('.cell-tasks');
        taskCell.innerHTML = '';
        const tasksNotesWrap = el('div', { className: 'tasks-notes-grid' });

        const tasksCol = el('div', { className: 'tn-col tasks-col' });
        tasksCol.appendChild(el('div', { className: 'tn-heading', textContent: 'Tasks' }));
        const tasksBody = el('div', { className: 'tn-body tasks-body' });
        if (p.tasks && p.tasks.length) {
            const renderTasks = (expanded) => {
                tasksBody.innerHTML = '';
                const tasksToShow = expanded ? p.tasks : p.tasks.slice(0, 2);
                tasksToShow.forEach(t => {
                    tasksBody.appendChild(el('div', {
                        className: `task-chip ${t.done ? 'done' : ''}`,
                        textContent: t.text || 'Task'
                    }));
                });
                if (p.tasks.length > 2) {
                    const moreBtn = el('button', {
                        className: 'btn-more-tasks',
                        textContent: expanded ? 'Show Less' : `+${p.tasks.length - 2} more`,
                        onclick: (e) => { e.stopPropagation(); renderTasks(!expanded); }
                    });
                    tasksBody.appendChild(moreBtn);
                }
            };
            renderTasks(false);
        } else {
            tasksBody.textContent = '--';
        }
        tasksCol.appendChild(tasksBody);

        const notesCol = el('div', { className: 'tn-col notes-col' });
        notesCol.appendChild(el('div', { className: 'tn-heading', textContent: 'Notes' }));
        const notesBody = el('div', { className: 'tn-body notes-body' });
        const notesText = (p.notes || '').trim();
        if (notesText) {
            notesBody.appendChild(el('div', { className: 'note-snippet', textContent: notesText }));
        } else {
            notesBody.textContent = '--';
        }
        notesCol.appendChild(notesBody);

        tasksNotesWrap.append(tasksCol, notesCol);
        taskCell.appendChild(tasksNotesWrap);

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
            className: `st status-btn st-${cls}`,
            type: 'button',
            textContent: label,
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
    [
        ['wait', 'Waiting'],
        ['work', 'Working'],
        ['pr', 'Pending Review'],
        ['comp', 'Complete'],
        ['del', 'Delivered']
    ].forEach(([cls, label]) => wrap.append(mk(cls, label)));
    return wrap;
}

function matches(q, p) {
    if (!q) return true;
    const str = (val) => (val || '').toLowerCase();
    return str(p.id).includes(q) || str(p.name).includes(q) || str(p.nick).includes(q) || str(p.notes).includes(q) || (p.tasks || []).some(t => str(t.text).includes(q)) || (p.statuses || []).some(s => str(s).includes(q));
}

function updateSortHeaders() {
    document.querySelectorAll('th[data-sort]').forEach(th => {
        th.classList.remove('sort-asc', 'sort-desc');
        if (th.dataset.sort === currentSort.key) th.classList.add(`sort-${currentSort.dir}`);
    });
}

// ===================== CRUD OPERATIONS =====================
function openEdit(i) {
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
    db.splice(i, 1);
    save();
    render();
}
function duplicate(i) {
    const original = db[i];
    const newProjectData = {
        id: original?.id || '',
        name: original?.name || '',
        path: original?.path || '',
        nick: '',
        notes: '',
        due: '',
        tasks: [],
        refs: [],
        statuses: []
    };
    editIndex = -1;
    document.getElementById('dlgTitle').textContent = 'Duplicate Project';
    document.getElementById('btnSaveProject').textContent = "Create Duplicate";
    fillForm(newProjectData);
    document.getElementById('editDlg').showModal();
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
    document.querySelectorAll('#taskList .task-row').forEach(row => {
        const text = row.querySelector('.t-text').value.trim();
        if (!text) return;
        const done = row.querySelector('.t-done').checked;
        const links = [row.querySelector('.t-link').value.trim(), row.querySelector('.t-link2').value.trim()].filter(Boolean).map(normalizeLink);
        out.tasks.push({ text, done, links });
    });
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
function setupStatusPicker(containerId, selected) {
    const elc = document.getElementById(containerId);
    const setPressed = () => elc.querySelectorAll('.st').forEach(b => b.setAttribute('aria-pressed', String(selected.includes(b.dataset.status))));
    if (!elc.__wired) {
        elc.addEventListener('click', e => {
            if (e.target.matches('.st')) {
                const s = e.target.dataset.status;
                const i = selected.indexOf(s);
                if (i >= 0) selected.splice(i, 1); else selected.push(s);
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
window.addTaskRow = () => addTaskRowFrom({});
window.addRefRow = () => addRefRowFrom({});
window.closeDlg = closeDlg;

// ===================== CSV & IMPORT/EXPORT =====================
function parseCSV(text) {
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
            tasks: [], refs: [], statuses: []
        };
        const parts = String(statusCell || '').split(/[,/|;]+/).map(s => s.trim()).filter(Boolean);
        for (const s of parts) {
            const c = canonStatus(s);
            if (c && !p.statuses.includes(c)) p.statuses.push(c);
        }
        const tparts = (tasksStr || '').replace(/\r/g, '\n').split(/\n|;|\u2022|\r/).map(s => s.trim()).filter(Boolean);
        for (const t of tparts) p.tasks.push({ text: t, done: false, links: [] });
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

// ===================== NOTES SYSTEM =====================
const debouncedSaveNotes = debounce(saveNotes, 500);

function renderNoteTabs() {
    const container = document.getElementById('notesTabsContainer');
    container.innerHTML = '';
    noteTabs.forEach(tabName => {
        const btn = el('button', {
            className: `inner-tab-btn ${tabName === activeNoteTab ? 'active' : ''}`,
            textContent: tabName,
            onclick: () => { activeNoteTab = tabName; renderNoteTabs(); renderNoteSearchResults(); }
        });
        const delIcon = el('span', {
            className: 'tab-delete-icon', textContent: '🗑️', title: 'Delete Page',
            onclick: (e) => {
                e.stopPropagation();
                if (confirm(`Permanently delete page "${tabName}"?`)) {
                    const idx = noteTabs.indexOf(tabName);
                    if (idx > -1) {
                        noteTabs.splice(idx, 1);
                        delete notesDb[tabName];
                        if (activeNoteTab === tabName) activeNoteTab = noteTabs.length > 0 ? noteTabs[Math.max(0, idx - 1)] : null;
                        saveNotes(); renderNoteTabs();
                    }
                }
            }
        });
        btn.appendChild(delIcon);
        container.appendChild(btn);
    });
    const addBtn = el('button', {
        className: 'add-tab-btn', textContent: '+', title: 'Add New Page',
        onclick: () => {
            const name = prompt('Enter name for new page:');
            if (name && name.trim()) {
                if (!noteTabs.includes(name.trim())) {
                    noteTabs.push(name.trim());
                    activeNoteTab = name.trim();
                    saveNotes(); renderNoteTabs();
                } else toast('Page name already exists.');
            }
        }
    });
    container.appendChild(addBtn);
    updateActiveNoteTextarea();
}

function updateActiveNoteTextarea() {
    const textarea = document.getElementById('notesTextarea');
    if (!activeNoteTab) {
        textarea.value = ''; textarea.placeholder = 'Create a page to begin.'; textarea.disabled = true; return;
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
    const query = val('notesSearch').toLowerCase();
    const resultsContainer = document.getElementById('notesSearchResults');
    resultsContainer.innerHTML = '';

    if (!query || !activeNoteTab) return;

    const queryWords = query.split(' ').filter(w => w);
    if (!queryWords.length) return;

    const content = notesDb[activeNoteTab];
    if (!content) return;

    const notes = content.split(/\n\s*\n/).filter(note => note.trim() !== '');

    notes.forEach((noteText) => {
        const lowerNoteText = noteText.toLowerCase();

        if (queryWords.every(word => lowerNoteText.includes(word))) {
            const item = el('div', { className: 'note-result-item', style: 'position: relative;' });
            const contentDiv = el('div', { className: 'note-result-content' });
            contentDiv.append(el('div', { className: 'snippet', textContent: noteText }));

            const copyIcon = el('button', { className: 'note-action-icon copy-icon', textContent: '📋', title: 'Copy' });
            copyIcon.onclick = () => { navigator.clipboard.writeText(noteText).then(() => toast('Copied!')); };

            const editIcon = el('button', { className: 'note-action-icon edit-icon', textContent: '✏️', title: 'Edit' });

            editIcon.onclick = () => {
                // 1. Scroll the main Page View to the top immediately
                window.scrollTo({ top: 0, behavior: 'smooth' });

                const textarea = document.getElementById('notesTextarea');
                const currentVal = textarea.value;
                const start = currentVal.indexOf(noteText);

                if (start !== -1) {
                    const end = start + noteText.length;

                    // 2. Focus and Select Text
                    textarea.blur();
                    textarea.setSelectionRange(start, end);
                    textarea.focus();

                    // 3. Calculate Exact Position using Mirror Div (Precision Scroll)
                    const mirror = document.createElement('div');
                    const style = window.getComputedStyle(textarea);

                    mirror.style.visibility = 'hidden';
                    mirror.style.position = 'absolute';
                    mirror.style.top = '-9999px';
                    mirror.style.whiteSpace = 'pre-wrap';
                    mirror.style.wordWrap = 'break-word';

                    ['fontFamily', 'fontSize', 'fontWeight', 'lineHeight', 'letterSpacing', 'padding'].forEach(p => {
                        mirror.style[p] = style[p];
                    });

                    mirror.style.boxSizing = 'border-box';
                    mirror.style.width = `${textarea.clientWidth}px`;
                    mirror.style.border = 'none';

                    mirror.textContent = currentVal.substring(0, start);

                    document.body.appendChild(mirror);
                    const targetY = mirror.clientHeight;
                    document.body.removeChild(mirror);

                    // 4. Scroll the Textarea internally to the specific line
                    textarea.scrollTop = Math.max(0, targetY - (textarea.clientHeight * 0.3));
                } else {
                    toast('Note content changed. Please refresh search.');
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
    const coreName = bundleName.replace('ElectricalCommands.', '').replace('.bundle', '');

    // Check Cache
    if (DESCRIPTION_CACHE[coreName]) return DESCRIPTION_CACHE[coreName];

    // Construct RAW URL. Pattern: AutoCADCommands/<CoreName>/<CoreName>_descriptions.json
    const url = `https://raw.githubusercontent.com/jacobhusband/ElectricalCommands/main/AutoCADCommands/${coreName}/${coreName}_descriptions.json`;

    try {
        const res = await fetch(url);
        if (!res.ok) throw new Error('Not found');
        const json = await res.json();
        DESCRIPTION_CACHE[coreName] = json; // Cache it
        return json;
    } catch (e) {
        console.warn(`Could not fetch description for ${coreName}`, e);
        return null;
    }
}

function openDetailsModal(name, descriptionData) {
    const dlg = document.getElementById('commandDetailsDlg');
    if (!dlg) return;

    document.getElementById('detailsTitle').textContent = name;
    const videoEl = document.getElementById('detailsVideo');
    videoEl.innerHTML = '';

    if (descriptionData && descriptionData.video && descriptionData.video.includes('loom.com')) {
        const videoId = descriptionData.video.split('/').pop();
        videoEl.append(el('iframe', { src: `https://www.loom.com/embed/${videoId}`, allowfullscreen: true }));
    } else {
        videoEl.innerHTML = '<p class="muted" style="padding:1rem;text-align:center">No video available.</p>';
    }

    const commandsEl = document.getElementById('detailsCommands');
    commandsEl.innerHTML = '';

    if (descriptionData && descriptionData.commands) {
        const list = el('ul', {}, Object.entries(descriptionData.commands).map(([cmd, desc]) =>
            el('li', {}, [el('strong', { textContent: cmd }), `: ${desc}`])
        ));
        commandsEl.append(el('div', { className: 'bundle-commands' }, [el('h4', { textContent: 'Commands' }), list]));
    } else {
        commandsEl.textContent = "No command details found.";
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

        // Process each bundle
        for (const bundle of response.data) {
            // Normalize name
            const coreName = bundle.name.replace('ElectricalCommands.', '').replace('.bundle', '');

            // Fetch description (background)
            const description = await fetchDescriptionForBundle(bundle.name);

            const card = el('div', { className: 'release-card' });
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

            const header = el('div', { className: 'release-card-header' }, [
                el('div', { className: 'release-card-title' }, [
                    el('div', { className: `bundle-status ${statusClass}`, title: statusTitle }),
                    el('span', { textContent: coreName })
                ]),
                el('button', {
                    className: 'info-btn', textContent: '?', title: 'Details',
                    onclick: () => openDetailsModal(coreName, description)
                })
            ]);

            const body = el('div', { className: 'release-card-body' });
            const tags = el('div', { className: 'command-tags' });

            if (description && description.commands) {
                Object.keys(description.commands).forEach(cmd =>
                    tags.append(el('span', { className: 'command-tag', textContent: cmd }))
                );
            }
            body.append(tags);

            const footer = el('div', { className: 'release-card-footer' });
            const btn = el('button', { className: `btn ${btnClass}`, textContent: btnText });
            btn.dataset.bundleName = bundle.bundle_name;
            btn.dataset.actionType = btnText;

            // Note: We rely on 'bundle.asset' from the backend for installation URL.
            // If backend doesn't provide it, we assume standard release naming.
            if (bundle.state !== 'installed' && bundle.asset) {
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

async function handleBundleAction(e) {
    const button = e.target.closest('[data-action-type]');
    if (!button) return;

    const actionType = button.dataset.actionType;
    button.disabled = true;
    button.textContent = 'Processing...';

    try {
        let response;
        if (actionType === 'Install' || actionType === 'Update') {
            if (!button.dataset.asset) throw new Error("Installation data missing.");
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
window.updateToolStatus = function (toolId, message) {
    const card = document.getElementById(toolId);
    if (!card) return;
    const statusEl = card.querySelector('.tool-card-status');
    const abortBtn = document.getElementById('abortBtn');
    statusEl.classList.remove('error');
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
    const mainTabContainer = document.querySelector('.main-nav');
    const notesResults = document.getElementById('notesSearchResults');

    mainTabContainer.addEventListener('click', e => {
        if (!e.target.matches('.main-tab-btn')) return;
        const tab = e.target.dataset.tab;

        document.querySelectorAll('.main-tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === tab));
        document.querySelectorAll('.tab-panel').forEach(p => {
            p.hidden = p.id !== `${tab}-panel`;
            p.classList.toggle('active', p.id === `${tab}-panel`);
        });

        if (notesResults) notesResults.innerHTML = '';

        if (tab === 'plugins') {
            loadAndRenderBundles();
        } else if (tab === 'projects') {
            render();
        }
    });
}

function initEventListeners() {
    document.getElementById('search').addEventListener('input', debounce(() => render(), 250));
    document.getElementById('notesSearch').addEventListener('input', debounce(() => renderNoteSearchResults(), 250));

    document.getElementById('quickNew').onclick = openNew;
    document.getElementById('settingsBtn').onclick = async () => {
        await populateSettingsModal();
        document.getElementById('settingsDlg').showModal();
    };
    document.getElementById('statsBtn').onclick = () => document.getElementById('statsDlg').showModal();
    document.getElementById('settings_howToSetupBtn').onclick = () => document.getElementById('apiKeyHelpDlg').showModal();
    document.getElementById('btnSaveProject').onclick = onSaveProject;

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

    document.getElementById('toolPublishDwgs').addEventListener('click', async (e) => {
        if (e.currentTarget.classList.contains('running')) return;
        if (!userSettings.autocadPath) {
            await showAutocadSelectModal();
            return;
        }
        e.currentTarget.classList.add('running');
        window.updateToolStatus('toolPublishDwgs', 'Initializing...');
        await window.pywebview.api.run_publish_script();
    });

    document.getElementById('toolCleanXrefs').addEventListener('click', async (e) => {
        if (e.currentTarget.classList.contains('running')) return;
        if (!userSettings.autocadPath) {
            await showAutocadSelectModal();
            return;
        }
        e.currentTarget.classList.add('running');
        window.updateToolStatus('toolCleanXrefs', 'Initializing...');
        await window.pywebview.api.run_clean_xrefs_script();
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
        if (!userSettings.autocadPath) {
            await showAutocadSelectModal();
            return;
        }
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

    document.getElementById('btnSelectTitleblock').addEventListener('click', async () => {
        const res = await window.pywebview.api.select_files({ allow_multiple: false, file_types: ['AutoCAD Drawings (*.dwg)'] });
        if (res.status === 'success' && res.paths.length) {
            document.getElementById('cleanDwgs_titleblockPath').value = res.paths[0];
        }
    });

    document.getElementById('btnBrowseAutocad').addEventListener('click', async () => {
        const res = await window.pywebview.api.select_files({
            allow_multiple: false,
            file_types: ['Executable Files (*.exe)'],
            directory: 'C:\\Program Files\\Autodesk'
        });
        if (res.status === 'success' && res.paths.length) {
            document.getElementById('settings_autocadPath').value = res.paths[0];
            // Uncheck radio buttons
            document.querySelectorAll('input[name="autocad_version_radio"]').forEach(radio => radio.checked = false);
        }
    });

    document.getElementById('btnBrowseAutocadSelect').addEventListener('click', async () => {
        const res = await window.pywebview.api.select_files({
            allow_multiple: false,
            file_types: ['Executable Files (*.exe)'],
            directory: 'C:\\Program Files\\Autodesk'
        });
        if (res.status === 'success' && res.paths.length) {
            document.getElementById('autocad_select_custom').value = res.paths[0];
            // Uncheck radio buttons
            document.querySelectorAll('input[name="autocad_select_radio"]').forEach(radio => radio.checked = false);
        }
    });

    document.getElementById('btnSaveAutocadSelect').addEventListener('click', async () => {
        const selectedRadio = document.querySelector('input[name="autocad_select_radio"]:checked');
        if (selectedRadio) {
            userSettings.autocadPath = selectedRadio.value;
        } else {
            userSettings.autocadPath = document.getElementById('autocad_select_custom').value.trim();
        }
        try {
            await window.pywebview.api.save_user_settings(userSettings);
        } catch (e) {
            toast('⚠️ Could not save settings.');
        }
        closeDlg('autocadSelectDlg');
    });

    const handleAI = () => {
        if (!userSettings.apiKey) { toast('Setup API Key in Settings first.'); return; }
        document.getElementById('emailArea').value = '';
        document.getElementById('aiSpinner').style.display = 'none';
        document.getElementById('emailDlg').showModal();
    };
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

    document.getElementById('notesTextarea').addEventListener('input', handleNoteInput);
    document.getElementById('settings_userName').oninput = (e) => { userSettings.userName = e.target.value; debouncedSaveUserSettings(); };
    document.getElementById('settings_apiKey').oninput = (e) => { userSettings.apiKey = e.target.value; debouncedSaveUserSettings(); };
    document.querySelectorAll('input[name="settings_discipline_radio"]').forEach(radio => {
        radio.onchange = (e) => {
            userSettings.discipline = e.target.value;
            debouncedSaveUserSettings();
        };
    });

    document.getElementById('btnPasteImport').onclick = () => {
        const rows = parseCSV(val('pasteArea'));
        importRows(rows, document.getElementById('hasHeader').checked);
        closeDlg('pasteDlg');
    };

    document.getElementById('btnMarkOverdue').onclick = () => {
        document.getElementById('markOverdueDlg').showModal();
    };
    document.getElementById('btnConfirmMarkOverdue').onclick = async () => {
        try {
            const response = await window.pywebview.api.mark_overdue_projects_complete();
            if (response.status === 'success') {
                toast(`Marked ${response.count} projects as complete.`);
                db = await load();
                render();
            } else {
                toast('Failed to mark projects as complete.');
            }
        } catch (e) {
            toast('Error marking projects as complete.');
        }
        closeDlg('markOverdueDlg');
    };
    document.getElementById('btnMarkOverdueDelivered').onclick = () => {
        document.getElementById('markOverdueDeliveredDlg').showModal();
    };
    document.getElementById('btnConfirmMarkOverdueDelivered').onclick = async () => {
        try {
            const response = await window.pywebview.api.mark_overdue_projects_delivered();
            if (response.status === 'success') {
                toast(`Marked ${response.count} projects as delivered.`);
                db = await load();
                render();
            } else {
                toast('Failed to mark projects as delivered.');
            }
        } catch (e) {
            toast('Error marking projects as delivered.');
        }
        closeDlg('markOverdueDeliveredDlg');
    };

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

    document.getElementById('btnDeleteAllNotes').onclick = () => {
        document.getElementById('deleteNotesConfirmInput').value = '';
        document.getElementById('btnDeleteNotesConfirm').disabled = true;
        document.getElementById('deleteNotesDlg').showModal();
    };
    document.getElementById('deleteNotesConfirmInput').oninput = (e) => {
        document.getElementById('btnDeleteNotesConfirm').disabled = e.target.value !== 'DELETE';
    };
    document.getElementById('btnDeleteNotesConfirm').onclick = async () => {
        try {
            const response = await window.pywebview.api.delete_all_notes();
            if (response.status === 'success') {
                toast('All notes data deleted.');
                notesDb = {};
                noteTabs = ['General'];
                activeNoteTab = 'General';
                renderNoteTabs();
            } else {
                toast('Failed to delete notes data.');
            }
        } catch (e) {
            toast('Error deleting notes data.');
        }
        closeDlg('deleteNotesDlg');
    };

    document.getElementById('btnUninstallAllPlugins').onclick = async () => {
        try {
            const response = await window.pywebview.api.check_autocad_running();
            if (response.is_running) {
                toast('AutoCAD is currently running. Please close AutoCAD and try again.');
                return;
            }
            document.getElementById('uninstallPluginsConfirmInput').value = '';
            document.getElementById('btnUninstallPluginsConfirm').disabled = true;
            document.getElementById('uninstallPluginsDlg').showModal();
        } catch (e) {
            toast('Error checking AutoCAD status.');
        }
    };
    document.getElementById('uninstallPluginsConfirmInput').oninput = (e) => {
        document.getElementById('btnUninstallPluginsConfirm').disabled = e.target.value !== 'UNINSTALL';
    };
    document.getElementById('btnUninstallPluginsConfirm').onclick = async () => {
        try {
            const response = await window.pywebview.api.uninstall_all_plugins();
            if (response.status === 'success') {
                toast(`Uninstalled ${response.count} plugins.`);
                loadAndRenderBundles();
            } else {
                toast('Failed to uninstall plugins.');
            }
        } catch (e) {
            toast('Error uninstalling plugins.');
        }
        closeDlg('uninstallPluginsDlg');
    };

    document.getElementById('commands-container').addEventListener('click', handleBundleAction);
    document.getElementById('btnInstallPrereq').onclick = async () => {
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
    if (!window.pywebview) await new Promise(r => window.addEventListener('pywebviewready', r));
    initEventListeners();
    initTabbedInterfaces();
    updateStickyOffsets();
    await loadUserSettings();
    db = await load();
    await loadNotes();
    renderNoteTabs();
    render();
    updateCleanDwgToolState();
}

init();
