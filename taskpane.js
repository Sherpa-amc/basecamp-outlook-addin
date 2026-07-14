/* Base Camp Outlook Add-in - taskpane.js
 *
 * Logic ported from the legacy VBA macro (FactsheetSaver). Calls the local
 * Flask server (https://localhost:8765) for matching, saving, and reading-pack
 * copying.
 */

const SERVER = 'https://localhost:8765';

// State for the per-attachment flow
let allSavedResults = [];
let pendingAttachments = [];
let currentAttachment = null;
let allFunds = [];   // populated once for the "change fund" picker

Office.onReady(() => {
    // Server health check on load
    checkServerHealth();

    // Preload the fund list (for the "change fund" picker).
    loadAllFunds();

    // Wire buttons
    document.getElementById('btn-process').onclick = startProcessing;
    document.getElementById('dt-ok').onclick = onDocTypeOk;
    document.getElementById('dt-skip').onclick = onDocTypeSkip;

    // Change-fund picker (doc-type dialog)
    document.getElementById('dt-change-fund').onclick = toggleFundPicker;
    document.getElementById('dt-fund-search').addEventListener('input', (e) => filterFundPicker(e.target.value));
    document.getElementById('dt-fund-search').addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            const first = document.querySelector('#dt-fund-dropdown .dropdown-item[data-id]');
            if (first) first.click();
        }
    });
    document.getElementById('nf-ok').onclick = onNewFundOk;
    document.getElementById('nf-skip').onclick = onNewFundSkip;
    document.getElementById('rp-ok').onclick = onReadingPackOk;
    document.getElementById('rp-skip').onclick = onReadingPackSkip;
    document.getElementById('rn-ok').onclick = onRenameOk;
    document.getElementById('rn-skip').onclick = onRenameSkip;
    document.getElementById('sum-close').onclick = resetUI;
    document.getElementById('err-back').onclick = resetUI;

    // Show attachment count summary
    showAttachmentSummary();
});

// =============================================================================
// Server health check
// =============================================================================

async function checkServerHealth() {
    setServerStatus('checking', 'Connecting...');
    try {
        const res = await fetch(`${SERVER}/api/outlook/health`);
        if (res.ok) {
            const data = await res.json();
            setServerStatus('ok', `Connected (v${data.version})`);
        } else {
            setServerStatus('error', `Server error ${res.status}`);
        }
    } catch (err) {
        setServerStatus('error', 'Server not running');
    }
}

function setServerStatus(state, text) {
    const el = document.getElementById('server-status');
    el.className = `srv srv-${state}`;
    document.getElementById('srv-text').textContent = text;

    // Disable process button if server is down
    const btn = document.getElementById('btn-process');
    btn.disabled = state === 'error';
    if (state === 'error') {
        btn.title = 'Local server not running. Start it via "Open Fund Manager.bat".';
    } else {
        btn.title = '';
    }
}

// =============================================================================
// Initial scan: list PDF attachments in the current email
// =============================================================================

function showAttachmentSummary() {
    const item = Office.context.mailbox.item;
    if (!item) {
        document.getElementById('att-summary').textContent = 'No email selected.';
        return;
    }
    const pdfs = (item.attachments || []).filter(isPdfAttachment);
    if (pdfs.length === 0) {
        document.getElementById('att-summary').textContent = 'No PDF attachments in this email.';
        document.getElementById('btn-process').disabled = true;
        return;
    }
    const word = pdfs.length === 1 ? 'attachment' : 'attachments';
    document.getElementById('att-summary').textContent =
        `${pdfs.length} PDF ${word} found.`;
}

function isPdfAttachment(att) {
    if (!att || !att.name) return false;
    if (att.attachmentType && att.attachmentType !== 'file') return false;
    return att.name.toLowerCase().endsWith('.pdf');
}

// =============================================================================
// Main processing flow (state machine)
// =============================================================================

async function startProcessing() {
    const item = Office.context.mailbox.item;
    if (!item) {
        showError('No email selected.');
        return;
    }

    const pdfs = (item.attachments || []).filter(isPdfAttachment);
    if (pdfs.length === 0) {
        showError('No PDF attachments to process.');
        return;
    }

    // Refresh server status
    await checkServerHealth();
    if (document.getElementById('server-status').classList.contains('srv-error')) {
        showError('Local server not running. Please start "Open Fund Manager.bat" and retry.');
        return;
    }

    pendingAttachments = pdfs.slice();
    allSavedResults = [];
    showSection('state-processing');
    processNextAttachment();
}

async function processNextAttachment() {
    if (pendingAttachments.length === 0) {
        showSummary();
        return;
    }

    currentAttachment = pendingAttachments.shift();
    setProcessingTitle(currentAttachment.name, 'Matching to fund...');

    const item = Office.context.mailbox.item;
    const subject = item.subject || '';
    const sender = (item.from && item.from.emailAddress) || '';

    // 1. Match
    let match;
    try {
        const res = await fetch(`${SERVER}/api/outlook/match-attachment`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                subject: subject,
                sender: sender,
                filename: currentAttachment.name,
            }),
        });
        if (!res.ok) throw new Error(`match failed: ${res.status}`);
        match = await res.json();
    } catch (err) {
        recordResult('error', currentAttachment.name, '(match failed)', err.message);
        return processNextAttachment();
    }

    if (match.matched) {
        currentAttachment._fund = match.fund;
        showDocTypePicker(currentAttachment.name);
    } else {
        showNewFundDialog(currentAttachment.name, match.suggested_name || '');
    }
}

// -----------------------------------------------------------------------------
// Doc type picker
// -----------------------------------------------------------------------------

function showDocTypePicker(filename) {
    document.getElementById('dt-filename').textContent = filename;
    document.querySelector('input[name="dt"][value="Factsheet"]').checked = true;

    // Show the matched fund and reset the (hidden) change-fund picker.
    document.getElementById('dt-fund-name').textContent =
        (currentAttachment._fund && currentAttachment._fund.name) || '(none)';
    document.getElementById('dt-fund-picker').classList.add('hidden');
    document.getElementById('dt-fund-dropdown').classList.add('hidden');
    document.getElementById('dt-fund-search').value = '';

    showSection('dialog-doctype');
}

// -----------------------------------------------------------------------------
// Change-fund picker (lets the user correct a wrong / missing match)
// -----------------------------------------------------------------------------

async function loadAllFunds() {
    try {
        const res = await fetch(`${SERVER}/api/funds`);
        if (!res.ok) return;
        const data = await res.json();
        allFunds = (data.funds || []).slice().sort((a, b) => a.name.localeCompare(b.name));
    } catch (e) {
        // Best-effort — the picker just won't have options if this fails.
    }
}

function toggleFundPicker() {
    const picker = document.getElementById('dt-fund-picker');
    const willShow = picker.classList.contains('hidden');
    picker.classList.toggle('hidden', !willShow);
    if (willShow) {
        const search = document.getElementById('dt-fund-search');
        search.value = '';
        filterFundPicker('');
        setTimeout(() => search.focus(), 50);
    }
}

function filterFundPicker(query) {
    const dropdown = document.getElementById('dt-fund-dropdown');
    const q = (query || '').toLowerCase().trim();

    let funds = allFunds;
    if (q) {
        funds = funds.filter(f =>
            f.name.toLowerCase().includes(q) ||
            (f.keywords || []).some(kw => (kw || '').includes(q))
        );
    }

    // Always offer a "create new fund" action at the top — the matched fund
    // may be wrong AND the right one may not exist yet.
    const typed = (query || '').trim();
    const createLabel = typed
        ? `+ Create new fund “${escHtml(typed)}”`
        : '+ Create new fund…';
    let html = `<div class="dropdown-item dropdown-create" data-create="1">` +
        `<span class="di-name">${createLabel}</span></div>`;

    if (funds.length === 0) {
        html += '<div class="dropdown-item dropdown-empty">No existing funds match.</div>';
    } else {
        html += funds.slice(0, 50).map(f =>
            `<div class="dropdown-item" data-id="${escAttr(f.id)}" data-name="${escAttr(f.name)}">` +
            `<span class="di-name">${escHtml(f.name)}</span>` +
            `<span class="di-cat">${escHtml(f.category || '')}</span></div>`
        ).join('');
    }
    dropdown.innerHTML = html;

    dropdown.querySelector('.dropdown-create').onclick = () => openNewFundFromPicker(typed);
    dropdown.querySelectorAll('.dropdown-item[data-id]').forEach(item => {
        item.onclick = () => pickFund(item.dataset.id, item.dataset.name);
    });
    dropdown.classList.remove('hidden');
}

// Open the create-fund dialog from the change-fund picker, seeding the name
// with whatever the user typed into the search box.
function openNewFundFromPicker(typedName) {
    document.getElementById('dt-fund-picker').classList.add('hidden');
    document.getElementById('dt-fund-dropdown').classList.add('hidden');
    document.getElementById('dt-fund-search').value = '';
    showNewFundDialog(currentAttachment.name, typedName || '');
}

function pickFund(id, name) {
    // Override the matched fund for the current attachment.
    const fund = allFunds.find(f => f.id === id) || { id, name };
    currentAttachment._fund = { id: fund.id, name: fund.name, category: fund.category || '' };
    document.getElementById('dt-fund-name').textContent = fund.name;
    document.getElementById('dt-fund-picker').classList.add('hidden');
    document.getElementById('dt-fund-dropdown').classList.add('hidden');
    document.getElementById('dt-fund-search').value = '';
}

function escHtml(s) {
    const d = document.createElement('div');
    d.textContent = s == null ? '' : String(s);
    return d.innerHTML;
}
function escAttr(s) {
    return escHtml(s).replace(/"/g, '&quot;');
}

async function onDocTypeOk() {
    const docType = document.querySelector('input[name="dt"]:checked').value;
    showSection('state-processing');
    setProcessingTitle(currentAttachment.name, 'Reading attachment bytes...');

    // 2. Get attachment bytes (base64)
    let contentBase64;
    try {
        contentBase64 = await getAttachmentContent(currentAttachment.id);
    } catch (err) {
        recordResult('error', currentAttachment.name,
            currentAttachment._fund.name, 'Could not read attachment: ' + err.message);
        return processNextAttachment();
    }

    // Stash so the rename flow can re-save the same bytes under a new name
    currentAttachment._docType = docType;
    currentAttachment._contentBase64 = contentBase64;

    performSave(docType, currentAttachment.name, contentBase64);
}

// 3. Save (shared by the doc-type flow and the rename retry)
async function performSave(docType, filename, contentBase64) {
    setProcessingTitle(currentAttachment.name, 'Saving to fund folder...');

    let saveRes;
    try {
        const res = await fetch(`${SERVER}/api/outlook/save-attachment`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                fund_id: currentAttachment._fund.id,
                doc_type: docType,
                filename: filename,
                content_base64: contentBase64,
            }),
        });

        if (res.status === 409) {
            // File already exists — offer to rename instead of skipping
            const err = await res.json().catch(() => ({}));
            return showRenameDialog(err.filename || filename);
        }

        if (!res.ok) {
            const err = await res.json().catch(() => ({}));
            throw new Error(err.error || err.message || `Server error ${res.status}`);
        }
        saveRes = await res.json();
    } catch (err) {
        recordResult('error', filename,
            currentAttachment._fund.name, 'Save failed: ' + err.message);
        return processNextAttachment();
    }

    recordResult('ok', filename,
        currentAttachment._fund.name, `${docType}: ${saveRes.subfolder}`);

    // 4. If Letter, prompt reading pack
    if (docType === 'Letter') {
        currentAttachment._savePath = saveRes.path;
        return showReadingPackDialog(filename);
    }

    processNextAttachment();
}

// -----------------------------------------------------------------------------
// Rename dialog (shown when the target filename already exists)
// -----------------------------------------------------------------------------

function showRenameDialog(existingName) {
    document.getElementById('rn-existing').textContent =
        `"${existingName}" already exists in this folder. Save under a new name:`;
    document.getElementById('rn-input').value = suggestRenamedFilename(existingName);
    showSection('dialog-rename');
    setTimeout(() => document.getElementById('rn-input').focus(), 50);
}

function onRenameOk() {
    const newName = document.getElementById('rn-input').value.trim();
    if (!newName) {
        alert('Please enter a file name.');
        return;
    }
    showSection('state-processing');
    // Re-save the stashed bytes. If the new name also exists, the rename
    // dialog re-appears prefilled with the latest attempted name.
    performSave(currentAttachment._docType, newName, currentAttachment._contentBase64);
}

function onRenameSkip() {
    recordResult('warn', currentAttachment.name,
        currentAttachment._fund.name, 'File already exists - skipped');
    showSection('state-processing');
    processNextAttachment();
}

// Append a " Mmm-yy" suffix before the extension, e.g. "Factsheet.pdf" ->
// "Factsheet Jun-26.pdf". The user can edit before saving.
function suggestRenamedFilename(filename) {
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                    'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const d = new Date();
    d.setDate(1);
    d.setMonth(d.getMonth() - 1);
    const suffix = ' ' + months[d.getMonth()] + '-' + String(d.getFullYear()).slice(-2);

    if (!filename) return 'document' + suffix + '.pdf';

    const dot = filename.lastIndexOf('.');
    if (dot <= 0) return filename + suffix;
    return filename.slice(0, dot) + suffix + filename.slice(dot);
}

function onDocTypeSkip() {
    recordResult('skip', currentAttachment.name,
        currentAttachment._fund.name, 'Skipped by user');
    showSection('state-processing');
    processNextAttachment();
}

// -----------------------------------------------------------------------------
// New fund dialog
// -----------------------------------------------------------------------------

function showNewFundDialog(filename, suggestedName) {
    document.getElementById('nf-filename').textContent = filename;
    document.getElementById('nf-name').value = suggestedName;
    document.getElementById('nf-keywords').value = suggestedName.toLowerCase();
    document.getElementById('nf-category').value = 'Universe';
    showSection('dialog-newfund');
    setTimeout(() => document.getElementById('nf-name').focus(), 50);
}

async function onNewFundOk() {
    const name = document.getElementById('nf-name').value.trim();
    if (!name) {
        alert('Please enter a fund name.');
        return;
    }
    const kwRaw = document.getElementById('nf-keywords').value.trim();
    const keywords = kwRaw
        ? kwRaw.split(',').map(s => s.trim().toLowerCase()).filter(Boolean)
        : [name.toLowerCase()];
    const category = document.getElementById('nf-category').value;

    showSection('state-processing');
    setProcessingTitle(currentAttachment.name, 'Creating fund...');

    // POST /api/funds (handles 409 duplicate-warning by retrying with skip flag)
    let res = await fetch(`${SERVER}/api/funds`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name, keywords, category }),
    });

    if (res.status === 409) {
        const data = await res.json().catch(() => ({}));
        const dupes = data.duplicates || [];
        const list = dupes.map(d => `  \u2022 ${d.name} (${d.category})`).join('\n');
        const proceed = confirm(
            `Possible duplicate(s) found:\n\n${list}\n\nCreate "${name}" anyway?`
        );
        if (!proceed) {
            recordResult('skip', currentAttachment.name, '(skipped)', 'Duplicate fund \u2014 not created');
            return processNextAttachment();
        }
        res = await fetch(`${SERVER}/api/funds`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ name, keywords, category, skip_duplicate_check: true }),
        });
    }

    if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        recordResult('error', currentAttachment.name, '(new fund)',
            'Create failed: ' + (err.error || err.message || res.status));
        return processNextAttachment();
    }

    const newFund = await res.json();
    currentAttachment._fund = newFund;

    showDocTypePicker(currentAttachment.name);
}

function onNewFundSkip() {
    recordResult('skip', currentAttachment.name, '(no fund)', 'Skipped by user');
    showSection('state-processing');
    processNextAttachment();
}

// -----------------------------------------------------------------------------
// Reading pack dialog (Letters only)
// -----------------------------------------------------------------------------

function showReadingPackDialog(filename) {
    document.getElementById('rp-filename').textContent = filename;
    document.querySelector('input[name="rp"][value="Manager"]').checked = true;
    showSection('dialog-readingpack');
}

async function onReadingPackOk() {
    const tag = document.querySelector('input[name="rp"]:checked').value;
    showSection('state-processing');
    setProcessingTitle(currentAttachment.name, `Copying to Reading Pack (${tag})...`);

    try {
        const res = await fetch(`${SERVER}/api/outlook/copy-to-reading-pack`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                source_path: currentAttachment._savePath,
                tag: tag,
            }),
        });
        if (!res.ok) {
            const err = await res.json().catch(() => ({}));
            throw new Error(err.error || err.message || `Server error ${res.status}`);
        }
        const result = await res.json();
        recordResult('ok', currentAttachment.name,
            currentAttachment._fund.name,
            `Reading Pack: ${tag}${result.duplicate ? ' (already there)' : ''}`);
    } catch (err) {
        recordResult('warn', currentAttachment.name,
            currentAttachment._fund.name,
            'Reading pack copy failed: ' + err.message);
    }

    processNextAttachment();
}

function onReadingPackSkip() {
    showSection('state-processing');
    processNextAttachment();
}

// =============================================================================
// Office.js helpers
// =============================================================================

function getAttachmentContent(attachmentId) {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.getAttachmentContentAsync(
            attachmentId,
            (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    // result.value.content is base64 (when format is Base64, the default)
                    resolve(result.value.content);
                } else {
                    reject(result.error || new Error('getAttachmentContentAsync failed'));
                }
            },
        );
    });
}

// =============================================================================
// UI helpers
// =============================================================================

const ALL_SECTIONS = [
    'state-idle',
    'state-processing',
    'dialog-doctype',
    'dialog-newfund',
    'dialog-rename',
    'dialog-readingpack',
    'state-summary',
    'state-error',
];

function showSection(id) {
    for (const sec of ALL_SECTIONS) {
        const el = document.getElementById(sec);
        if (el) el.classList.toggle('hidden', sec !== id);
    }
}

function setProcessingTitle(filename, step) {
    document.getElementById('proc-title').textContent = 'Processing';
    document.getElementById('proc-filename').textContent = filename || '';
    document.getElementById('proc-step').textContent = step || '';
}

function recordResult(status, filename, fundName, info) {
    allSavedResults.push({ status, filename, fundName, info });
}

function showSummary() {
    const ul = document.getElementById('sum-list');
    ul.innerHTML = '';
    let okCount = 0, skipCount = 0, errCount = 0, warnCount = 0;
    for (const r of allSavedResults) {
        const li = document.createElement('li');
        li.className = `sum-${r.status}`;
        li.textContent = `${r.fundName} \u2014 ${r.filename} (${r.info})`;
        ul.appendChild(li);
        if (r.status === 'ok') okCount++;
        else if (r.status === 'skip') skipCount++;
        else if (r.status === 'warn') warnCount++;
        else errCount++;
    }
    const parts = [];
    if (okCount) parts.push(`${okCount} saved`);
    if (warnCount) parts.push(`${warnCount} duplicates`);
    if (skipCount) parts.push(`${skipCount} skipped`);
    if (errCount) parts.push(`${errCount} errors`);
    document.getElementById('sum-title').textContent = parts.join(', ') || 'Done';
    showSection('state-summary');
}

function showError(message) {
    document.getElementById('err-message').textContent = message;
    showSection('state-error');
}

function resetUI() {
    pendingAttachments = [];
    allSavedResults = [];
    currentAttachment = null;
    showSection('state-idle');
    showAttachmentSummary();
}
