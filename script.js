// filepath: C:\Users\JacobH\Documents\dev\ProjectManagement\script.js
// ===================== STATE MANAGEMENT =====================
/** @typedef {{ label:string, url:string, raw?:string }} Link */
/** @typedef {{ text:string, done?:boolean, links:Link[] }} Task */
/** @typedef {{ id:string, name:string, nick?:string, notes?:string, due?:string, path?:string, tasks:Task[], refs:Link[], statuses?:string[], statusTags?:string[], status?:string }} Project */
let db = [];
let notesDb = { keyed: {}, general: {} };
let noteTabs = [];
let editIndex = -1;
let currentSort = { key: 'due', dir: 'desc' };
let statusFilter = 'all';
let dueFilter = 'all';
const STATUS_CANON = ["Pending Review", "Complete", "Waiting"];
const LABEL_TO_KEY = { 'Pending Review': 'pendingReview', 'Complete': 'complete', 'Waiting': 'waiting' };
const KEY_TO_LABEL = { pendingReview: 'Pending Review', complete: 'Complete', waiting: 'Waiting' };
let userSettings = {
    userName: '',
    discipline: 'Electrical', // Default value
    apiKey: ''
};
let chatHistory = [];
// ===================== SERVER I/O (PYWEBVIEW) =====================
async function load() {
    try {
        const arr = await window.pywebview.api.get_tasks();
        migrateStatuses(arr);
        return arr;
    } catch (e) {
        console.warn('Failed to load from Python backend.', e);
        toast('Could not load data.');
        return [];
    }
}
async function save() {
    try {
        const response = await window.pywebview.api.save_tasks(db);
        if (response.status !== 'success') throw new Error(response.message || 'Unknown error');
    } catch (e) {
        console.warn('Failed to save to Python backend.', e);
        toast('⚠️ Failed to save data.');
    }
}
async function loadNotes() {
    try {
        const data = await window.pywebview.api.get_notes() || {};
        // Expected data structure: { tabs: string[], keyed: object, general: object }
        noteTabs = Array.isArray(data.tabs) && data.tabs.length > 0 ? data.tabs : ['Default Tab'];
        notesDb = {
            keyed: data.keyed || {},
            general: data.general || {}
        };
        activeNoteTab = noteTabs[0];
        return notesDb;
    } catch (e) {
        console.warn('Failed to load notes from Python backend.', e);
        toast('Could not load notes.');
        noteTabs = ['Default Tab'];
        activeNoteTab = noteTabs[0];
        return {};
    }
}
async function saveNotes() {
    try {
        const dataToSave = {
            tabs: noteTabs,
            keyed: notesDb.keyed,
            general: notesDb.general
        };
        const response = await window.pywebview.api.save_notes(dataToSave);
        if (response.status !== 'success') throw new Error(response.message || 'Unknown error');
    } catch (e) {
        console.warn('Failed to save notes to Python backend.', e);
        toast('⚠️ Failed to save notes.');
    }
}
// ===================== UTILITIES & HELPERS =====================
function el(tag, props = {}, children = []) { const n = document.createElement(tag); Object.entries(props).forEach(([k, v]) => { if (k.startsWith('aria-') || k.startsWith('data-')) { if (v != null) n.setAttribute(k, String(v)); } else { n[k] = v; } }); children.forEach(c => n.append(c)); return n; }
function parseDueStr(s) { if (!s) return null; s = s.trim(); let m = s.match(/^\d{4}-\d{2}-\d{2}$/); if (m) return new Date(s + 'T12:00:00'); s = s.replace(/[.]/g, '/').replace(/\s+/g, ''); const parts = s.split('/'); if (parts.length === 3) { let [mm, dd, yy] = parts; if (yy.length === 2) { yy = '20' + yy; } const iso = `${yy}-${mm.padStart(2, '0')}-${dd.padStart(2, '0')}T12:00:00`; const d = new Date(iso); if (!isNaN(d)) return d; } const d2 = new Date(s); return isNaN(d2) ? null : d2; }
function dueState(dueStr) { const d = parseDueStr(dueStr); if (!d) return 'ok'; const today = new Date(); today.setHours(0, 0, 0, 0); const dayOfWeek = today.getDay(); const startOfWeek = new Date(today); startOfWeek.setDate(today.getDate() - dayOfWeek); const endOfWeek = new Date(startOfWeek); endOfWeek.setDate(startOfWeek.getDate() + 6); endOfWeek.setHours(23, 59, 59, 999); if (d >= startOfWeek && d <= endOfWeek) { return 'dueSoon'; } if (d < today) { return 'overdue'; } return 'ok'; }
function humanDate(s) { const d = parseDueStr(s); if (!d) return ''; return d.toLocaleDateString(undefined, { year: '2-digit', month: '2-digit', day: '2-digit' }); }
function basename(path) { try { if (!path) return ''; const norm = path.replace(/\\/g, '/'); const idx = norm.lastIndexOf('/'); return idx >= 0 ? norm.slice(idx + 1) : norm; } catch { return path; } }
function toFileURL(raw) { if (!raw) return ''; let s = raw.trim(); if (/^https?:\/\//i.test(s)) return s; if (/^\\\\/.test(s)) return 'file:' + s.replace(/^\\\\/, '/////').replace(/\\/g, '/'); if (/^[A-Za-z]:\\/.test(s)) return 'file:///' + s.replace(/\\/g, '/'); return s; }
function normalizeLink(input) { const raw = (input || '').trim(); const url = toFileURL(raw); const label = basename(raw) || raw || 'link'; return { label, url, raw }; }
function convertPath(raw) { if (raw.startsWith('\\\\acies.lan\\cachedrive\\projects2\\')) return raw.replace('\\\\acies.lan\\cachedrive\\projects2\\', 'M:\\'); if (raw.startsWith('\\\\acies.lan\\cachedrive\\projects\\')) return raw.replace('\\\\acies.lan\\cachedrive\\projects\\', 'P:\\'); return raw; }
function toast(msg, duration = 2200) { const t = el('div', { textContent: msg, style: 'position:fixed;bottom:16px;left:50%;transform:translateX(-50%);background:var(--panel);border:1px solid var(--border);padding:.5rem .75rem;border-radius:10px;z-index:9999;box-shadow:0 4px 12px rgba(0,0,0,0.2)' }); document.body.append(t); setTimeout(() => t.remove(), duration); }
function closeDlg(id) { document.getElementById(id).close(); }
function val(id) { return document.getElementById(id).value.trim(); }
const debounce = (fn, delay) => { let timeoutId; return (...args) => { clearTimeout(timeoutId); timeoutId = setTimeout(() => fn(...args), delay); }; };
// ===================== USER SETTINGS =====================
async function saveUserSettings() {
    try {
        await window.pywebview.api.save_user_settings(userSettings);
    } catch (e) {
        console.error("Failed to save user settings to backend:", e);
        toast('⚠️ Could not save settings.');
    }
}
async function loadUserSettings() {
    try {
        const storedSettings = await window.pywebview.api.get_user_settings();
        if (storedSettings) {
            userSettings = storedSettings;
        }
    } catch (e) {
        console.error("Failed to load user settings from backend:", e);
        userSettings = { userName: '', discipline: 'Electrical', apiKey: '' };
    }
}
const debouncedSaveUserSettings = debounce(saveUserSettings, 500);
// ===================== STATUS MIGRATION & HELPERS =====================
function canonStatus(s) { if (!s) return null; const t = String(s).trim().toLowerCase(); if (['pending review', 'pending-review', 'review', 'pr'].includes(t)) return 'Pending Review'; if (['complete', 'completed', 'done'].includes(t)) return 'Complete'; if (['waiting', 'wait'].includes(t)) return 'Waiting'; return null; }
function hasStatus(p, s) { return Array.isArray(p.statuses) && p.statuses.includes(s); }
function toggleStatus(p, label) { if (!Array.isArray(p.statuses)) p.statuses = []; const isOn = p.statuses.includes(label); const key = LABEL_TO_KEY[label]; if (key) setTag(p, key, !isOn); }
function syncStatusArrays(p) { if (!Array.isArray(p.statuses)) p.statuses = []; const fromTags = Array.isArray(p.statusTags) ? p.statusTags : []; for (const k of fromTags) { const L = KEY_TO_LABEL[k]; if (L && !p.statuses.includes(L)) p.statuses.push(L); } p.statuses = [...new Set(p.statuses.filter(s => STATUS_CANON.includes(s)))]; p.statusTags = p.statuses.map(s => LABEL_TO_KEY[s]).filter(Boolean); if (p.statuses.includes('Complete')) p.status = 'Complete'; else if (p.statuses.includes('Waiting')) p.status = 'Waiting'; else if (p.statuses.includes('Pending Review')) p.status = 'Pending'; else p.status = p.status || ''; }
function migrateStatuses(arr) { for (const p of arr) { if (!Array.isArray(p.statuses)) p.statuses = []; if (p.status) { String(p.status).split(/[,/|;]+/).map(s => s.trim()).filter(Boolean).forEach(piece => { const c = canonStatus(piece); if (c && !p.statuses.includes(c)) p.statuses.push(c); }); } p.statuses = p.statuses.filter(s => STATUS_CANON.includes(s)); syncStatusArrays(p); } }
function setTag(p, key, on) { const tags = getStatusTags(p); const idx = tags.indexOf(key); if (on && idx === -1) tags.push(key); if (!on && idx !== -1) tags.splice(idx, 1); p.statusTags = tags; const label = KEY_TO_LABEL[key]; if (!Array.isArray(p.statuses)) p.statuses = []; const j = p.statuses.indexOf(label); if (label) { if (on && j === -1) p.statuses.push(label); if (!on && j !== -1) p.statuses.splice(j, 1); } if (p.statuses.includes('Complete')) p.status = 'Complete'; else if (p.statuses.includes('Waiting')) p.status = 'Waiting'; else if (p.statuses.includes('Pending Review')) p.status = 'Pending'; else p.status = ''; }
function getStatusTags(p) { let tags = Array.isArray(p.statusTags) ? [...p.statusTags] : []; const s = (p.status || '').toLowerCase(); if (s) { if (s.includes('complete') && !tags.includes('complete')) tags.push('complete'); if (s.includes('waiting') && !tags.includes('waiting')) tags.push('waiting'); if ((s.includes('pending review') || s === 'pending') && !tags.includes('pendingReview')) tags.push('pendingReview'); } return [...new Set(tags)]; }
// ===================== RENDER =====================
function render() {
    const tbody = document.getElementById('tbody');
    const emptyState = document.getElementById('emptyState');
    tbody.innerHTML = '';
    const q = val('search').toLowerCase();
    let items = db.filter(p => {
        if (q && !matches(q, p)) return false;
        // Due date filters
        if (dueFilter === 'overdue' && (dueState(p.due) !== 'overdue' || hasStatus(p, 'Complete'))) return false;
        if (dueFilter === 'soon' && dueState(p.due) !== 'dueSoon') return false;
        if (dueFilter === 'ok' && dueState(p.due) !== 'ok') return false;
        // Status filters
        if (statusFilter === 'incomplete') {
            if (hasStatus(p, 'Complete')) return false;
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
            if (!da && !dbb) comparison = 0; else if (!da) comparison = 1; else if (!dbb) comparison = -1; else comparison = da - dbb;
        } else {
            comparison = String(valA || '').localeCompare(String(valB || ''), undefined, { numeric: true });
        }
        return comparison * (currentSort.dir === 'asc' ? 1 : -1);
    });
    updateSortHeaders();
    document.getElementById('kWeek').textContent = db.filter(p => dueState(p.due) === 'dueSoon' && !hasStatus(p, 'Complete')).length;
    emptyState.style.display = items.length ? 'none' : 'block';
    const rowTemplate = document.getElementById('project-row-template');
    items.forEach(p => {
        const tr = rowTemplate.content.cloneNode(true).querySelector('tr');
        const idx = db.indexOf(p);
        tr.querySelector('.cell-id').textContent = p.id || '—';
        const nameCell = tr.querySelector('.cell-name');
        if (p.path) {
            const link = el('button', { className: 'path-link', textContent: p.name || '—', title: `Open: ${p.path}` });
            link.onclick = async () => { try { await window.pywebview.api.open_path(convertPath(p.path)); toast('Opened in File Explorer'); } catch (e) { toast('Failed to open path.'); } };
            nameCell.append(link);
        } else {
            nameCell.textContent = p.name || '—';
        }
        if (p.nick) nameCell.append(el('small', { className: 'muted', textContent: ` (${p.nick})` }));
        const dueCell = tr.querySelector('.cell-due');
        if (p.due) {
            const ds = dueState(p.due);
            const pillClass = ds === 'overdue' ? 'pill overdue' : ds === 'dueSoon' ? 'pill dueSoon' : 'pill ok';
            dueCell.append(el('div', { className: pillClass, textContent: humanDate(p.due) }));
        } else {
            dueCell.textContent = '—';
        }
        tr.querySelector('.cell-status').append(renderStatusToggles(p));
        const taskCell = tr.querySelector('.cell-tasks');
        if (p.tasks?.length) {
            p.tasks.forEach(t => taskCell.append(el('div', { className: `task-chip ${t.done ? 'done' : ''}`, textContent: t.text || '—' })));
        } else {
            taskCell.textContent = '—';
        }
        const actionsCell = tr.querySelector('.cell-actions');
        actionsCell.append(
            el('button', { className: 'btn tiny', textContent: 'Edit', onclick: () => openEdit(idx) }),
            el('button', { className: 'btn tiny btn-danger', textContent: 'Del', onclick: () => removeProject(idx) }),
            el('button', { className: 'btn tiny', textContent: 'Dup', onclick: () => duplicate(idx) }),
        );
        tbody.append(tr);
    });
}
function renderStatusToggles(p) {
    const wrap = el('div', { className: 'status-group' });
    const mk = (cls, label) => {
        const b = el('button', { className: `st st-${cls}`, type: 'button', textContent: label.replace(' ', '\n'), title: label, 'aria-pressed': String(hasStatus(p, label)) });
        b.onclick = async (e) => { e.stopPropagation(); toggleStatus(p, label); await save(); render(); };
        return b;
    };
    wrap.append(mk('pr', 'Pending Review'), mk('comp', 'Complete'), mk('wait', 'Waiting'));
    return wrap;
}
function matches(q, p) { if (!q) return true; return ['id', 'name', 'nick', 'notes'].some(k => (p[k] || '').toLowerCase().includes(q)) || (p.tasks || []).some(t => (t.text || '').toLowerCase().includes(q)) || (p.statuses || []).some(s => s.toLowerCase().includes(q)); }
function updateSortHeaders() { document.querySelectorAll('th[data-sort]').forEach(th => { th.classList.remove('sort-asc', 'sort-desc'); if (th.dataset.sort === currentSort.key) th.classList.add(`sort-${currentSort.dir}`); }); }
// ===================== CRUD & ACTIONS =====================
function openEdit(i) { if (typeof i !== 'number' || i < 0 || !db[i]) return toast('Could not find row to edit.'); editIndex = i; const p = db[i]; document.getElementById('dlgTitle').textContent = `Edit Project — ${p.id || ''}`; fillForm(p); document.getElementById('editDlg').showModal(); }
function openNew() { editIndex = -1; document.getElementById('dlgTitle').textContent = 'New Project'; fillForm({ tasks: [], refs: [], statuses: [] }); document.getElementById('editDlg').showModal(); }
function removeProject(i) { if (!confirm('Delete this project?')) return; const id = db[i]?.id; db.splice(i, 1); save(); render(); toast(`Deleted ${id || 'project'}`); }
function duplicate(i) {
    const original = db[i];
    const newProjectData = {
        // Persistent info from original project
        id: original.id,
        name: original.name,
        nick: original.nick,
        path: original.path,
        refs: JSON.parse(JSON.stringify(original.refs || [])), // Keep refs, deep copy

        // Reset fields that are likely to change for the new assignment
        due: '',
        notes: '',
        tasks: [],
        statuses: []
    };

    editIndex = -1; // Ensure we are in "new project" mode
    document.getElementById('dlgTitle').textContent = 'Create Duplicate Project';
    fillForm(newProjectData);
    document.getElementById('editDlg').showModal();
}
function markOverdueAsComplete() { let count = 0; db.forEach(p => { if (dueState(p.due) === 'overdue') { setTag(p, 'complete', true); count++; } }); if (count > 0) { save(); render(); toast(`Marked ${count} overdue projects as complete`); } else { toast('No overdue projects found'); } }
// ===================== FORM HANDLING (EDIT/NEW MODAL) =====================
function fillForm(p) {
    document.getElementById('f_id').value = p.id || '';
    document.getElementById('f_name').value = p.name || '';
    document.getElementById('f_nick').value = p.nick || '';
    document.getElementById('f_notes').value = p.notes || '';
    document.getElementById('f_due').value = p.due || '';
    document.getElementById('f_path').value = p.path || '';
    setupStatusPicker('f_statuses', p.statuses || []);
    document.getElementById('taskList').innerHTML = '';
    const tasks = (p.tasks || []).map(task => typeof task === 'string' ? { text: task } : task);
    tasks.forEach(addTaskRowFrom);
    document.getElementById('refList').innerHTML = '';
    (p.refs || []).forEach(addRefRowFrom);
}
function readForm() {
    const out = { id: val('f_id'), name: val('f_name'), nick: val('f_nick'), notes: val('f_notes'), due: val('f_due'), path: val('f_path'), tasks: [], refs: [], statuses: readStatusPicker('f_statuses') };
    document.querySelectorAll('#taskList .task-row').forEach(row => { const text = row.querySelector('.t-text').value.trim(); if (!text) return; const done = row.querySelector('.t-done').checked; const links = [row.querySelector('.t-link').value.trim(), row.querySelector('.t-link2').value.trim()].filter(Boolean).map(normalizeLink); out.tasks.push({ text, done, links }); });
    document.querySelectorAll('#refList .ref-row').forEach(row => { const label = row.querySelector('.r-label').value.trim(); const raw = row.querySelector('.r-url').value.trim(); if (!raw) return; const link = normalizeLink(raw); if (label) link.label = label; out.refs.push(link); });
    syncStatusArrays(out);
    return out;
}
function onSaveProject() { const data = readForm(); if (editIndex >= 0) { db[editIndex] = data; toast('Project updated'); } else { db.push(data); toast('Project added'); } save(); render(); closeDlg('editDlg'); }
function setupStatusPicker(containerId, selected) { const elc = document.getElementById(containerId); const setPressed = () => elc.querySelectorAll('.st').forEach(b => b.setAttribute('aria-pressed', String(selected.includes(b.dataset.status)))); if (!elc.__wired) { elc.addEventListener('click', e => { if (e.target.matches('.st')) { const s = e.target.dataset.status; const i = selected.indexOf(s); if (i >= 0) selected.splice(i, 1); else selected.push(s); setPressed(); } }); elc.__wired = true; } setPressed(); }
function readStatusPicker(containerId) { return Array.from(document.querySelectorAll(`#${containerId} .st[aria-pressed="true"]`)).map(b => b.dataset.status); }
function addTaskRow() { addTaskRowFrom({}); }
function addTaskRowFrom(t = {}) { const template = document.getElementById('task-row-template'); const row = template.content.cloneNode(true).querySelector('.task-row'); row.querySelector('.t-text').value = t.text || ''; row.querySelector('.t-done').checked = !!t.done; row.querySelector('.t-link').value = t.links?.[0]?.raw || ''; row.querySelector('.t-link2').value = t.links?.[1]?.raw || ''; row.querySelector('.btn-remove').onclick = () => row.remove(); document.getElementById('taskList').append(row); }
function addRefRow() { addRefRowFrom({}); }
function addRefRowFrom(L = {}) { const template = document.getElementById('ref-row-template'); const row = template.content.cloneNode(true).querySelector('.ref-row'); row.querySelector('.r-label').value = L.label || ''; row.querySelector('.r-url').value = L.raw || L.url || ''; row.querySelector('.btn-remove').onclick = () => row.remove(); document.getElementById('refList').append(row); }
// ===================== CSV Import/Export =====================
function parseCSV(text) { const rows = []; let i = 0, field = '', row = [], inQ = false; function pushField() { row.push(field); field = ''; } function pushRow() { rows.push(row); row = []; } while (i < text.length) { const ch = text[i]; if (inQ) { if (ch === '"') { if (text[i + 1] === '"') { field += '"'; i += 2; continue; } inQ = false; i++; continue; } field += ch; i++; continue; } else { if (ch === '"') { inQ = true; i++; continue; } if (ch === ',') { pushField(); i++; continue; } if (ch === '\n') { pushField(); pushRow(); i++; continue; } if (ch === '\r') { if (text[i + 1] === '\n') { i += 2; pushField(); pushRow(); } else { i++; pushField(); pushRow(); } continue; } field += ch; i++; } } if (field !== '' || row.length) { pushField(); pushRow(); } return rows; }
function importRows(rows, hasHeader = true) { if (rows.length && !hasHeader) { const joined = rows[0].map(s => (s || '').toUpperCase()).join(' | '); if (joined.includes('PROJECT NAME') || joined.includes('DUE')) hasHeader = true; } if (hasHeader) rows = rows.slice(1); let added = 0; for (const r of rows) { if (!r.length) continue; const [id, name, nick, notes, due, statusCell, tasksStr, path, ...refs] = r; if (!(id || name || tasksStr || refs.some(Boolean))) continue; const p = { id: String(id || '').trim(), name: (name || '').trim(), nick: (nick || '').trim(), notes: (notes || '').trim(), due: (due || '').trim(), path: (path || '').trim(), tasks: [], refs: [], statuses: [] }; const parts = String(statusCell || '').split(/[,/|;]+/).map(s => s.trim()).filter(Boolean); for (const s of parts) { const c = canonStatus(s); if (c && !p.statuses.includes(c)) p.statuses.push(c); } const tparts = (tasksStr || '').replace(/\r/g, '\n').split(/\n|;|\u2022|\r/).map(s => s.trim()).filter(Boolean); for (const t of tparts) p.tasks.push({ text: t, done: false, links: [] }); for (const cell of refs) { const s = (cell || '').trim(); if (!s) continue; p.refs.push(normalizeLink(s)); } db.push(p); added++; } save(); render(); toast(`Imported ${added} rows`); }
function exportCSV() { const header = ['ID', 'PROJECT NAME', 'NICKNAME', 'NOTES', 'DUE DATE', 'STATUS', 'TASKS', 'PATH', 'REFERENCE1', 'REFERENCE2', 'REFERENCE3', 'REFERENCE4']; const lines = [header.join(',')]; for (const p of db) { const tasks = (p.tasks || []).map(t => t.text).join(' | '); const refs = (p.refs || []).map(L => L.raw || L.url).slice(0, 4); const statusStr = (p.statuses || []).join(' | '); const row = [p.id, p.name, p.nick || '', (p.notes || '').replaceAll('\n', ' '), p.due || '', statusStr, tasks, p.path || '', ...refs]; const csv = row.map(cell => { const s = String(cell ?? ''); return /[",\n]/.test(s) ? '"' + s.replaceAll('"', '""') + '"' : s; }).join(','); lines.push(csv); } const blob = new Blob([lines.join('\n')], { type: 'text/csv' }); const url = URL.createObjectURL(blob); const a = el('a', { href: url, download: 'projects.csv' }); document.body.append(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 1000); }
function exportJSON() { const blob = new Blob([JSON.stringify(db, null, 2)], { type: 'application/json' }); const url = URL.createObjectURL(blob); const a = el('a', { href: url, download: 'projects.json' }); document.body.append(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 1000); }
// ===================== NOTES MANAGEMENT =====================
const debouncedSaveNotes = debounce(saveNotes, 500);
let activeNoteCategory = 'keyed';
let activeNoteTab = null;

function renderNoteTabs() {
    const container = document.getElementById('notesTabsContainer');
    container.innerHTML = '';
    noteTabs.forEach(tabName => {
        const btn = el('button', {
            className: `inner-tab-btn ${tabName === activeNoteTab ? 'active' : ''}`,
            textContent: tabName,
            'data-plan-tab': tabName
        });
        container.append(btn);
    });
    updateActiveNoteTextarea();
}

function updateActiveNoteTextarea() {
    const textarea = document.getElementById('notesTextarea');
    if (!activeNoteTab) {
        textarea.value = '';
        textarea.placeholder = 'Create a tab to begin.';
        textarea.disabled = true;
        return;
    }
    textarea.disabled = false;
    textarea.placeholder = `Enter ${activeNoteCategory} notes for ${activeNoteTab}...`;
    textarea.value = notesDb[activeNoteCategory]?.[activeNoteTab] || '';
}

function handleNoteInput(e) {
    if (!activeNoteCategory || !activeNoteTab) return;
    if (!notesDb[activeNoteCategory]) {
        notesDb[activeNoteCategory] = {};
    }
    notesDb[activeNoteCategory][activeNoteTab] = e.target.value;
    debouncedSaveNotes();
}

function deleteNoteParagraph(category, plan, index) {
    if (!confirm('Are you sure you want to delete this note paragraph?')) return;
    const fullContent = notesDb[category]?.[plan] || '';
    const notes = fullContent.split(/\n\s*\n/).filter(note => note.trim() !== '');
    notes.splice(index, 1);
    const newContent = notes.join('\n\n');
    if (!notesDb[category]) notesDb[category] = {};
    notesDb[category][plan] = newContent;
    saveNotes(); // Save immediately
    renderNoteTabs(); // Re-render to update textarea
    renderNoteSearchResults(); // Re-run search
}

function renderNoteSearchResults() {
    const query = val('search').toLowerCase();
    const resultsContainer = document.getElementById('notesSearchResults');
    resultsContainer.innerHTML = '';
    if (!query) return;

    const queryWords = query.split(' ').filter(w => w);
    if (!queryWords.length) return;

    const CATEGORY_LABELS = { keyed: 'Keyed Notes', general: 'General Notes' };
    const seenNotes = new Set(); // Track unique notes to avoid duplicates

    for (const category in notesDb) {
        for (const plan in notesDb[category]) {
            const normalizedPlan = plan.toLowerCase(); // Normalize tab name for comparison
            const content = notesDb[category][plan];
            if (!content) continue;

            const notes = content.split(/\n\s*\n/).filter(note => note.trim() !== '');
            notes.forEach((noteText, index) => {
                const lowerNoteText = noteText.toLowerCase();
                if (queryWords.every(word => lowerNoteText.includes(word))) {
                    // Create a unique key for the note based on category, plan, and content
                    const noteKey = `${category}:${normalizedPlan}:${noteText}`;
                    if (seenNotes.has(noteKey)) return; // Skip if already processed
                    seenNotes.add(noteKey);

                    // Use the original plan name (not normalized) for display
                    const location = `${CATEGORY_LABELS[category]} - ${plan}`;
                    const item = el('div', { className: 'note-result-item' });
                    const contentDiv = el('div', { className: 'note-result-content' });
                    const actionsDiv = el('div', { className: 'note-result-actions' });

                    contentDiv.append(
                        el('div', { className: 'location', textContent: location }),
                        el('div', { className: 'snippet', textContent: noteText })
                    );

                    const editBtn = el('button', { className: 'btn tiny', textContent: 'Edit' });
                    editBtn.onclick = () => {
                        activeNoteCategory = category;
                        activeNoteTab = plan; // Use original plan name
                        document.querySelectorAll('[data-category-toggle]').forEach(btn => btn.classList.toggle('active', btn.dataset.categoryToggle === category));
                        renderNoteTabs(); // Re-render tabs to show the active one
                        const searchInput = document.getElementById('search');
                        searchInput.value = '';
                        renderNoteSearchResults();
                        setTimeout(() => {
                            const targetTextarea = document.getElementById('notesTextarea');
                            if (targetTextarea) {
                                targetTextarea.focus();
                                const pos = targetTextarea.value.indexOf(noteText);
                                if (pos !== -1) {
                                    targetTextarea.setSelectionRange(pos, pos + noteText.length);
                                    const textBefore = targetTextarea.value.substring(0, pos);
                                    const lineBreaks = (textBefore.match(/\n/g) || []).length;
                                    const lineHeight = targetTextarea.scrollHeight / targetTextarea.value.split('\n').length;
                                    targetTextarea.scrollTop = lineBreaks * lineHeight;
                                }
                            }
                        }, 50);
                    };

                    const copyBtn = el('button', { className: 'btn tiny', textContent: 'Copy' });
                    copyBtn.onclick = () => {
                        navigator.clipboard.writeText(noteText)
                            .then(() => toast('Note copied to clipboard!'))
                            .catch(err => toast('Failed to copy note.'));
                    };

                    const deleteBtn = el('button', { className: 'btn tiny btn-danger', textContent: 'Delete' });
                    deleteBtn.onclick = () => deleteNoteParagraph(category, plan, index);

                    actionsDiv.append(editBtn, copyBtn, deleteBtn);
                    item.append(contentDiv, actionsDiv);
                    resultsContainer.append(item);
                }
            });
        }
    }
}
// ===================== TAB MANAGEMENT =====================
function handleMainSearch() {
    const activeTab = document.querySelector('.main-tab-btn.active')?.dataset.tab || 'projects';
    if (activeTab === 'notes') {
        renderNoteSearchResults();
    } else {
        render();
    }
}

function initTabbedInterfaces() {
    // Main tabs
    const mainTabContainer = document.querySelector('.main-nav');
    const mainSearchInput = document.getElementById('search');

    mainTabContainer.addEventListener('click', e => {
        if (!e.target.matches('.main-tab-btn')) return;
        const targetTab = e.target.dataset.tab;
        mainTabContainer.querySelectorAll('.main-tab-btn').forEach(btn => btn.classList.toggle('active', btn.dataset.tab === targetTab));
        document.querySelectorAll('.tab-panel').forEach(panel => {
            panel.hidden = panel.id !== `${targetTab}-panel`;
            panel.classList.toggle('active', panel.id === `${targetTab}-panel`);
        });

        if (targetTab === 'notes') {
            mainSearchInput.placeholder = 'Search notes...';
            mainSearchInput.disabled = false;
        } else if (targetTab === 'chat') {
            mainSearchInput.placeholder = 'Search disabled';
            mainSearchInput.disabled = true;
        } else { // This covers 'projects' and 'tools'
            mainSearchInput.placeholder = 'Search projects...';
            mainSearchInput.disabled = false;
        }
        // Clear search and re-render for the new context
        mainSearchInput.value = '';
        handleMainSearch();
    });
    // Notes Category Toggle
    const categoryToggle = document.querySelector('.notes-category-toggle');
    categoryToggle.addEventListener('click', e => {
        if (!e.target.matches('[data-category-toggle]')) return;
        activeNoteCategory = e.target.dataset.categoryToggle;
        categoryToggle.querySelectorAll('[data-category-toggle]').forEach(btn => btn.classList.remove('active'));
        e.target.classList.add('active');
        updateActiveNoteTextarea();
    });
    // Inner tabs (for notes section)
    const innerTabContainer = document.getElementById('notesTabsContainer');
    innerTabContainer.addEventListener('click', e => {
        if (!e.target.matches('.inner-tab-btn')) return;
        activeNoteTab = e.target.dataset.planTab;
        renderNoteTabs();
    });
}
// ===================== TOOL STATUS UPDATER =====================
// This function is called by the Python backend
window.updateToolStatus = function (toolId, message) {
    const card = document.getElementById(toolId);
    if (!card) return;

    const statusEl = card.querySelector('.tool-card-status');
    statusEl.classList.remove('error');

    if (message.startsWith('ERROR:')) {
        const errorMsg = message.substring(6).trim();
        statusEl.textContent = `Error: ${errorMsg}`;
        statusEl.classList.add('error');
        card.classList.add('running'); // Keep it in running state to show error
        // Optionally, reset after a delay
        setTimeout(() => {
            card.classList.remove('running');
        }, 5000);
    } else if (message === 'DONE') {
        statusEl.textContent = 'Completed successfully!';
        setTimeout(() => {
            card.classList.remove('running');
        }, 2000); // Keep success message for 2 seconds
    } else {
        statusEl.textContent = message;
    }
}
// ===================== EVENT WIRING =====================
function openSettingsModal() {
    document.getElementById('settings_userName').value = userSettings.userName || '';
    document.getElementById('settings_apiKey').value = userSettings.apiKey || '';
    const disciplineRadios = document.querySelectorAll('#settings_discipline input[name="settings_discipline_radio"]');
    disciplineRadios.forEach(radio => {
        radio.checked = (radio.value === userSettings.discipline);
    });
    document.getElementById('settingsDlg').showModal();
}

function initEventListeners() {

    console.log('Verifying critical elements...');
    const criticalElements = ['btnProceedCleanDwgs', 'toolCleanDwgs', 'cleanDwgs_titleblockPath'];
    criticalElements.forEach(id => {
        if (!document.getElementById(id)) {
            console.error(`❌ CRITICAL: Element '${id}' not found!`);
        }
    });

    document.getElementById('quickNew').addEventListener('click', openNew);
    document.getElementById('settingsBtn').addEventListener('click', openSettingsModal);
    document.getElementById('menuBtn').addEventListener('click', openDrawer);
    document.getElementById('drawerClose').addEventListener('click', closeDrawer);
    document.getElementById('backdrop').addEventListener('click', closeDrawer);
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && document.getElementById('drawer').classList.contains('open')) closeDrawer(); });
    document.getElementById('btnNew').addEventListener('click', openNew);
    document.getElementById('btnImport').addEventListener('click', async () => { try { const result = await window.pywebview.api.import_and_process_csv(); if (result.status === 'success') { if (result.data?.length > 0) { db.push(...result.data); migrateStatuses(db); await save(); render(); toast(`Successfully imported ${result.data.length} projects.`); } else { toast('No new projects were found in the selected file.'); } } else if (result.status !== 'cancelled') { throw new Error(result.message || 'An unknown error occurred during import.'); } } catch (e) { console.error('CSV Import failed:', e); toast(`⚠️ Import failed: ${e.message}`); } });
    document.getElementById('btnPaste').addEventListener('click', () => document.getElementById('pasteDlg').showModal());
    document.getElementById('btnExportCsv').addEventListener('click', exportCSV);
    document.getElementById('btnExportJson').addEventListener('click', exportJSON);
    document.getElementById('btnMarkOverdue').addEventListener('click', markOverdueAsComplete);
    document.getElementById('search').addEventListener('input', debounce(handleMainSearch, 250));
    document.getElementById('dueFilterGroup').addEventListener('click', (e) => { if (e.target.tagName === 'BUTTON') { dueFilter = e.target.dataset.dueFilter; document.querySelectorAll('#dueFilterGroup .btn').forEach(b => b.classList.remove('active')); e.target.classList.add('active'); render(); } });
    document.getElementById('statusFilterGroup').addEventListener('click', (e) => { if (e.target.tagName === 'BUTTON') { statusFilter = e.target.dataset.filter; document.querySelectorAll('#statusFilterGroup .btn').forEach(b => b.classList.remove('active')); e.target.classList.add('active'); render(); } });
    document.querySelector('.table thead').addEventListener('click', e => { const th = e.target.closest('th[data-sort]'); if (!th) return; const key = th.dataset.sort; if (currentSort.key === key) { currentSort.dir = currentSort.dir === 'asc' ? 'desc' : 'asc'; } else { currentSort.key = key; currentSort.dir = key === 'due' ? 'desc' : 'asc'; } render(); });
    document.getElementById('btnPasteImport').addEventListener('click', () => { const txt = val('pasteArea'); const rows = parseCSV(txt); const hasHeader = document.getElementById('hasHeader').checked; importRows(rows, hasHeader); closeDlg('pasteDlg'); document.getElementById('pasteArea').value = ''; });
    document.getElementById('btnSaveProject').addEventListener('click', onSaveProject);
    document.getElementById('btnCreateFolder').addEventListener('click', async () => { const path = val('f_path'); if (!path) return toast('Project path is empty.'); try { const res = await window.pywebview.api.create_folder(path); if (res.status === 'success') toast('Folder created successfully.'); else throw new Error(res.message); } catch (e) { toast(`⚠️ Error creating folder: ${e.message}`); } });
    document.getElementById('btnDeleteAll').addEventListener('click', () => { document.getElementById('deleteConfirmInput').value = ''; document.getElementById('btnDeleteConfirm').disabled = true; document.getElementById('deleteDlg').showModal(); });
    document.getElementById('deleteConfirmInput').addEventListener('input', (e) => { document.getElementById('btnDeleteConfirm').disabled = e.target.value !== 'DELETE'; });
    document.getElementById('btnDeleteConfirm').addEventListener('click', () => { if (val('deleteConfirmInput') === 'DELETE') { db = []; save(); render(); closeDlg('deleteDlg'); toast('All project data has been deleted.'); } });

    const openAiModal = () => {
        if (!userSettings.apiKey) {
            toast('Please set up your Gemini API key in the settings first.');
            openSettingsModal();
            return;
        }
        document.getElementById('emailArea').value = '';
        document.getElementById('aiSpinner').style.display = 'none';
        document.getElementById('emailDlg').showModal();
    };
    document.getElementById('btnEmail').addEventListener('click', openAiModal);
    document.getElementById('aiBtn').addEventListener('click', openAiModal);
    document.getElementById('settings_howToSetupBtn').addEventListener('click', () => document.getElementById('apiKeyHelpDlg').showModal());

    // Settings Dialog listeners
    document.getElementById('settings_userName').addEventListener('input', (e) => {
        userSettings.userName = e.target.value;
        debouncedSaveUserSettings();
    });
    document.getElementById('settings_apiKey').addEventListener('input', (e) => {
        userSettings.apiKey = e.target.value;
        debouncedSaveUserSettings();
    });
    document.getElementById('settings_discipline').addEventListener('change', (e) => {
        if (e.target.name === 'settings_discipline_radio') {
            userSettings.discipline = e.target.value;
            debouncedSaveUserSettings();
        }
    });

    document.getElementById('btnProcessEmail').addEventListener('click', async () => {
        const emailText = val('emailArea');
        if (!emailText) return toast('Please paste email content first.');
        if (!userSettings.apiKey) {
            toast('Please set your Gemini API key in the settings (⚙️).');
            openSettingsModal();
            return;
        }
        const spinner = document.getElementById('aiSpinner');
        spinner.style.display = 'block';
        try {
            const result = await window.pywebview.api.process_email_with_ai(
                emailText,
                userSettings.apiKey,
                userSettings.userName,
                userSettings.discipline
            );
            if (result.status === 'success') {
                closeDlg('emailDlg');
                openNew();
                fillForm(result.data);
                toast('AI analysis complete. Please review and save.');
            } else {
                throw new Error(result.message || 'An unknown AI error occurred.');
            }
        } catch (e) {
            console.error('AI Processing failed:', e);
            toast(`⚠️ AI Error: ${e.message}`, 5000);
        } finally {
            spinner.style.display = 'none';
        }
    });
    // Notes functionality
    document.getElementById('notesTextarea').addEventListener('input', handleNoteInput);
    document.getElementById('addNoteTabBtn').addEventListener('click', () => {
        const newTabName = prompt('Enter a name for the new tab:');
        if (newTabName && !noteTabs.includes(newTabName)) {
            noteTabs.push(newTabName);
            activeNoteTab = newTabName;
            saveNotes();
            renderNoteTabs();
        } else if (newTabName) {
            toast('A tab with that name already exists.');
        }
    });
    document.getElementById('removeNoteTabBtn').addEventListener('click', () => {
        if (!activeNoteTab) return toast('No active tab to remove.');
        if (!confirm(`Are you sure you want to permanently delete the "${activeNoteTab}" tab and all its notes?`)) return;

        const index = noteTabs.indexOf(activeNoteTab);
        if (index > -1) {
            noteTabs.splice(index, 1);
            delete notesDb.keyed[activeNoteTab];
            delete notesDb.general[activeNoteTab];

            if (noteTabs.length === 0) {
                activeNoteTab = null;
            } else {
                activeNoteTab = noteTabs[Math.max(0, index - 1)];
            }
            saveNotes();
            renderNoteTabs();
        }
    });
    // Tools Tab
    document.getElementById('toolPublishDwgs').addEventListener('click', async (e) => {
        const card = e.currentTarget;
        if (card.classList.contains('running')) return;
        card.classList.add('running');
        window.updateToolStatus('toolPublishDwgs', 'Starting...');
        try {
            await window.pywebview.api.run_publish_script();
        } catch (e) {
            window.updateToolStatus('toolPublishDwgs', `ERROR: ${e.message}`);
        }
    });
    document.getElementById('toolCleanXrefs').addEventListener('click', async (e) => {
        const card = e.currentTarget;
        if (card.classList.contains('running')) return;
        card.classList.add('running');
        window.updateToolStatus('toolCleanXrefs', 'Starting...');
        try {
            await window.pywebview.api.run_clean_xrefs_script();
        } catch (e) {
            window.updateToolStatus('toolCleanXrefs', `ERROR: ${e.message}`);
        }
    });
    document.getElementById('toolCleanDwgs').addEventListener('click', (e) => {
        const card = e.currentTarget;
        if (card.classList.contains('running')) return;
        // Reset form
        document.getElementById('cleanDwgs_titleblockPath').value = '';
        document.querySelector('input[name="cleanDwgs_size_radio"][value="22x34"]').checked = true;
        // Show dialog
        document.getElementById('cleanDwgsDlg').showModal();
    });

    // Clean DWGs Dialog - NULL-SAFE
    const btnSelectTitleblock = document.getElementById('btnSelectTitleblock');
    if (btnSelectTitleblock) {
        btnSelectTitleblock.addEventListener('click', async () => {
            try {
                const result = await window.pywebview.api.select_files({
                    allow_multiple: false,
                    file_types: ['AutoCAD Drawings (*.dwg)']
                });
                if (result.status === 'success' && result.paths.length > 0) {
                    const titleblockPathEl = document.getElementById('cleanDwgs_titleblockPath');
                    if (titleblockPathEl) titleblockPathEl.value = result.paths[0];
                }
            } catch (e) {
                toast(`⚠️ Error selecting file: ${e.message}`);
            }
        });
    } else {
        console.error('❌ btnSelectTitleblock not found!');
    }

    // CRITICAL: Add null check before adding event listener
    const btnProceedCleanDwgs = document.getElementById('btnProceedCleanDwgs');
    if (btnProceedCleanDwgs) {
        btnProceedCleanDwgs.addEventListener('click', async () => {
            const titleblockPath = val('cleanDwgs_titleblockPath');
            const sizeEl = document.querySelector('input[name="cleanDwgs_size_radio"]:checked');
            const size = sizeEl ? sizeEl.value : '22x34';

            // Get selected disciplines (titleblock parent is automatic)
            const selectedDisciplines = Array.from(document.querySelectorAll('input[name="cleanDwgs_discipline"]:checked'))
                .map(cb => cb.value);

            if (!titleblockPath) {
                toast('Please select a titleblock DWG.');
                return;
            }

            const data = {
                titleblock: titleblockPath,
                disciplines: selectedDisciplines,
                size: size
            };

            closeDlg('cleanDwgsDlg');
            const card = document.getElementById('toolCleanDwgs');
            if (card) card.classList.add('running');
            window.updateToolStatus('toolCleanDwgs', 'Starting...');

            try {
                await window.pywebview.api.run_clean_dwgs_script(data);
            } catch (e) {
                window.updateToolStatus('toolCleanDwgs', `ERROR: ${e.message}`);
            }
        });
    } else {
        console.error('❌ CRITICAL: btnProceedCleanDwgs not found!');
    }

    // Chat Tab
    document.getElementById('chat-send-btn').addEventListener('click', handleSendMessage);
    document.getElementById('chat-input').addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && e.ctrlKey) {
            e.preventDefault(); // prevent new line
            handleSendMessage();
        }
    });
    document.getElementById('newChatBtn').addEventListener('click', startNewChat);

    // Initialize tab functionality
    initTabbedInterfaces();
}
// ===================== CHAT FUNCTIONS =====================
/**
 * Appends a message to the chat container and updates the code file buttons.
 * @param {string} html The message content (can be HTML).
 * @param {'user' | 'bot'} sender The sender of the message.
 */
function appendChatMessage(html, sender) {
    const messagesContainer = document.getElementById('chat-messages');
    const codeFilesContainer = document.getElementById('chat-code-files-container');
    const messageType = sender === 'user' ? 'user-message' : 'bot-message';

    const messageDiv = el('div', { className: `chat-message ${messageType}` });
    messageDiv.innerHTML = html; // Use innerHTML to render parsed markdown/code blocks

    messagesContainer.append(messageDiv);
    messagesContainer.scrollTop = messagesContainer.scrollHeight; // Auto-scroll to bottom

    // Add event listeners for any new copy buttons inside the message
    messageDiv.querySelectorAll('.copy-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const code = btn.closest('.code-block').querySelector('code').innerText;
            navigator.clipboard.writeText(code).then(() => {
                btn.textContent = 'Copied!';
                setTimeout(() => { btn.textContent = 'Copy'; }, 2000);
            }).catch(err => {
                toast('Failed to copy code.');
                console.error('Copy failed', err);
            });
        });
    });

    // If it's a bot message, detect code blocks and create buttons below the chat window
    if (sender === 'bot') {
        messageDiv.querySelectorAll('.code-block').forEach((codeBlock, index) => {
            const language = codeBlock.querySelector('.code-block-header span')?.textContent || 'code';
            const code = codeBlock.querySelector('code')?.innerText;

            if (!code) return;

            const fileButton = el('div', { className: 'code-file-button' });
            const fileNameSpan = el('span', { className: 'file-name', textContent: `${language}_${index + 1}` });
            const copyBtn = el('button', { className: 'copy-btn', textContent: 'Copy' });

            copyBtn.addEventListener('click', () => {
                navigator.clipboard.writeText(code).then(() => {
                    copyBtn.textContent = 'Copied!';
                    toast(`Copied ${language} code block.`);
                    setTimeout(() => { copyBtn.textContent = 'Copy'; }, 2000);
                }).catch(err => {
                    toast('Failed to copy code.');
                    console.error('Copy failed', err);
                });
            });

            fileButton.append(fileNameSpan, copyBtn);
            codeFilesContainer.append(fileButton);
        });
    }
}

/**
 * Parses the bot's response, separating text from code blocks.
 * @param {string} responseText The raw text from the Gemini API.
 * @returns {string} HTML string with formatted code blocks.
 */
function parseBotResponse(responseText) {
    const codeBlockRegex = /```(\w*)\n([\s\S]*?)\n```/g;
    let lastIndex = 0;
    let html = '';

    let match;
    while ((match = codeBlockRegex.exec(responseText)) !== null) {
        // Add the text before the code block
        const precedingText = responseText.substring(lastIndex, match.index).trim();
        if (precedingText) {
            html += `<p>${precedingText.replace(/\n/g, '<br>')}</p>`;
        }

        const language = match[1] || 'code';
        const code = match[2].trim();

        // Add the formatted code block
        html += `
            <div class="code-block">
                <div class="code-block-header">
                    <span>${language}</span>
                    <button class="copy-btn">Copy</button>
                </div>
                <pre><code>${code.replace(/</g, "<").replace(/>/g, ">")}</code></pre>
            </div>
        `;

        lastIndex = codeBlockRegex.lastIndex;
    }

    // Add any remaining text after the last code block
    const remainingText = responseText.substring(lastIndex).trim();
    if (remainingText) {
        html += `<p>${remainingText.replace(/\n/g, '<br>')}</p>`;
    }

    return html;
}

function startNewChat() {
    if (confirm('Are you sure you want to start a new chat? Your current conversation will be lost.')) {
        chatHistory = [];
        document.getElementById('chat-messages').innerHTML = '';
        document.getElementById('chat-code-files-container').innerHTML = '';
        appendChatMessage('<p>New chat started. How can I help you?</p>', 'bot');
    }
}

async function handleSendMessage() {
    const input = document.getElementById('chat-input');
    const message = input.value.trim();
    if (!message) return;

    if (!userSettings.apiKey) {
        toast('Please set up your Gemini API key in the settings (⚙️) first.');
        openSettingsModal();
        return;
    }

    // Display user's message and add to history
    appendChatMessage(message.replace(/</g, "<").replace(/>/g, ">"), 'user');
    chatHistory.push({ role: 'user', parts: [{ text: message }] });
    input.value = '';
    input.focus();

    // Show thinking indicator
    const thinkingMsg = el('div', { className: 'chat-message bot-message thinking' });
    thinkingMsg.innerHTML = '<div class="spinner">Thinking...</div>';
    const messagesContainer = document.getElementById('chat-messages');
    messagesContainer.append(thinkingMsg);
    messagesContainer.scrollTop = messagesContainer.scrollHeight;

    try {
        const result = await window.pywebview.api.get_chat_response(chatHistory);
        thinkingMsg.remove();

        if (result.status === 'success') {
            const formattedHtml = parseBotResponse(result.response);
            appendChatMessage(formattedHtml, 'bot');
            chatHistory.push({ role: 'model', parts: [{ text: result.response }] });
        } else {
            throw new Error(result.response);
        }
    } catch (e) {
        thinkingMsg.remove();
        const errorHtml = `<strong>Error:</strong> ${e.message}`;
        appendChatMessage(errorHtml, 'bot');
        console.error('Chat error:', e);
    }
}
// ===================== INITIALIZATION =====================
async function init() {
    // Wait for the webview Python backend to be ready first.
    await new Promise(resolve => window.addEventListener('pywebviewready', resolve));

    try {
        // Now that the API is ready, set up all event listeners.
        initEventListeners();

        // Load all data from the backend.
        await loadUserSettings();
        db = await load();
        await loadNotes();

        // Render the UI with the loaded data.
        renderNoteTabs();
        render();
    } catch (error) {
        console.error("Fatal error during application initialization:", error);
        document.body.innerHTML = `<div style="padding: 2rem; text-align: center; color: #ef4444;">
            <h1>Application Failed to Start</h1>
            <p>Could not load initial data. Please check the console for errors (Right-click > Inspect) and restart the application.</p>
            <p><b>Error:</b> ${error.message}</p>
        </div>`;
    }
}

// Start the application
init();

function openDrawer() { document.getElementById('drawer').classList.add('open'); document.getElementById('backdrop').hidden = false; document.getElementById('menuBtn')?.setAttribute('aria-expanded', 'true'); }
function closeDrawer() { document.getElementById('drawer').classList.remove('open'); document.getElementById('backdrop').hidden = true; document.getElementById('menuBtn')?.setAttribute('aria-expanded', 'false'); }
