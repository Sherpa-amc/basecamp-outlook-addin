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

Office.onReady(() => {
    // Server health check on load
    checkServerHealth();

    // Wire buttons
    document.getElementById('btn-process').onclick = startProcessing;
    document.getElementById('dt-ok').onclick = onDocTypeOk;
    document.getElementById('dt-skip').onclick = onDocTypeSkip;
    document.getElementById('nf-ok').onclick = onNewFundOk;
    document.getElementById('nf-skip').onclick = onNewFundSkip;
    document.getElementById('rp-ok').onclick = onReadingPackOk;
    document.getElementById('rp-skip').onclick = onReadingPackSkip;
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
    showSection('dialog-doctype');
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

    setProcessingTitle(currentAttachment.name, 'Saving to fund folder...');

    // 3. Save
    let saveRes;
    try {
        const res = await fetch(`${SERVER}/api/outlook/save-attachment`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                fund_id: currentAttachment._fund.id,
                doc_type: docType,
                filename: currentAttachment.name,
                content_base64: contentBase64,
            }),
        });

        if (res.status === 409) {
            const err = await res.json().catch(() => ({}));
            recordResult('warn', currentAttachment.name,
                currentAttachment._fund.name, 'File already exists \u2014 not saved again');
            return processNextAttachment();
        }

        if (!res.ok) {
            const err = await res.json().catch(() => ({}));
            throw new Error(err.error || err.message || `Server error ${res.status}`);
        }
        saveRes = await res.json();
    } catch (err) {
        recordResult('error', currentAttachment.name,
            currentAttachment._fund.name, 'Save failed: ' + err.message);
        return processNextAttachment();
    }

    recordResult('ok', currentAttachment.name,
        currentAttachment._fund.name, `${docType}: ${saveRes.subfolder}`);

    // 4. If Letter, prompt reading pack
    if (docType === 'Letter') {
        currentAttachment._savePath = saveRes.path;
        return showReadingPackDialog(currentAttachment.name);
    }

    processNextAttachment();
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
