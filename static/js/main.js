/* ============================================================
   main.js — Corporate Tools Suite Frontend Logic
   ============================================================ */

// ─── Toast Notifications ──────────────────────────────────────
function showToast(message, type = 'success', duration = 4000) {
  const container = document.getElementById('toast-container');
  if (!container) return;

  const icons = { success: '✅', error: '❌', warning: '⚠️', info: 'ℹ️' };
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  toast.innerHTML = `
    <span class="toast-icon">${icons[type] || 'ℹ️'}</span>
    <span class="toast-msg">${message}</span>
  `;
  container.appendChild(toast);
  setTimeout(() => {
    toast.style.animation = 'fadeOut .3s ease forwards';
    setTimeout(() => toast.remove(), 300);
  }, duration);
}

// ─── Sidebar Toggle ───────────────────────────────────────────
function initSidebar() {
  const toggle = document.getElementById('menu-toggle');
  const sidebar = document.getElementById('sidebar');
  const overlay = document.getElementById('sidebar-overlay');

  if (toggle) {
    toggle.addEventListener('click', () => {
      sidebar.classList.toggle('open');
      overlay.classList.toggle('show');
    });
  }
  if (overlay) {
    overlay.addEventListener('click', () => {
      sidebar.classList.remove('open');
      overlay.classList.remove('show');
    });
  }

  // Highlight active nav item
  const links = document.querySelectorAll('.nav-item');
  const path = window.location.pathname;
  links.forEach(link => {
    const href = link.getAttribute('href') || '';
    if (path === href || (href !== '/' && path.startsWith(href))) {
      link.classList.add('active');
    }
  });
}

// ─── Tabs ─────────────────────────────────────────────────────
function initTabs(containerId) {
  const container = document.getElementById(containerId) || document;
  const tabBtns = container.querySelectorAll('.tab-btn');
  const tabContents = container.querySelectorAll('.tab-content');

  tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      const target = btn.dataset.tab;
      tabBtns.forEach(b => b.classList.remove('active'));
      tabContents.forEach(c => c.classList.remove('active'));
      btn.classList.add('active');
      const content = document.getElementById(target);
      if (content) content.classList.add('active');
    });
  });
}

// ─── Drag & Drop Upload Zone ──────────────────────────────────
function initDropZone(zoneId, inputId, onFilesSelected) {
  const zone = document.getElementById(zoneId);
  const input = document.getElementById(inputId);
  if (!zone) return;

  zone.addEventListener('dragover', e => {
    e.preventDefault();
    zone.classList.add('drag-over');
  });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.classList.remove('drag-over');
    const files = Array.from(e.dataTransfer.files);
    if (onFilesSelected) onFilesSelected(files);
  });
  if (input) {
    input.addEventListener('change', () => {
      const files = Array.from(input.files);
      if (onFilesSelected) onFilesSelected(files);
    });
  }
}

// ─── Format Bytes ─────────────────────────────────────────────
function formatBytes(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(2) + ' MB';
}

// ─── File Extension Badge ─────────────────────────────────────
function getExtBadge(filename) {
  const ext = filename.split('.').pop().toLowerCase();
  const badges = {
    pdf: 'ext-pdf', xlsx: 'ext-excel', xls: 'ext-excel',
    csv: 'ext-csv', doc: 'ext-word', docx: 'ext-word',
    jpg: 'ext-image', jpeg: 'ext-image', png: 'ext-image',
    gif: 'ext-image', zip: 'ext-zip', rar: 'ext-zip'
  };
  return { ext: ext.toUpperCase(), cls: badges[ext] || 'ext-file' };
}

// ─── Render File Item ─────────────────────────────────────────
function renderFileItem(file, actions = '') {
  const { ext, cls } = getExtBadge(file.name || file);
  const size = file.size ? formatBytes(file.size) : '';
  return `
    <div class="file-item" id="file-${Date.now()}-${Math.random().toString(36).substr(2,5)}">
      <div class="file-ext-badge ${cls}">${ext}</div>
      <div class="file-info">
        <div class="file-name">${file.name || file}</div>
        ${size ? `<div class="file-meta">${size}</div>` : ''}
      </div>
      <div class="file-actions">${actions}</div>
    </div>`;
}

// ─── Upload Files via Fetch ───────────────────────────────────
async function uploadFiles(url, formData, progressEl, resultEl) {
  if (progressEl) {
    progressEl.style.display = 'block';
    const bar = progressEl.querySelector('.progress-bar');
    if (bar) { bar.style.width = '0%'; animateProgress(bar, 80, 1500); }
  }

  try {
    const res = await fetch(url, { method: 'POST', body: formData });
    const data = await res.json();

    if (progressEl) {
      const bar = progressEl.querySelector('.progress-bar');
      if (bar) bar.style.width = '100%';
      setTimeout(() => { progressEl.style.display = 'none'; }, 600);
    }
    return data;
  } catch (err) {
    if (progressEl) progressEl.style.display = 'none';
    return { success: false, error: err.message };
  }
}

function animateProgress(bar, targetPct, duration) {
  const start = performance.now();
  function step(now) {
    const pct = Math.min(((now - start) / duration) * targetPct, targetPct);
    bar.style.width = pct + '%';
    if (pct < targetPct) requestAnimationFrame(step);
  }
  requestAnimationFrame(step);
}

// ─── Show Result Box ──────────────────────────────────────────
function showResult(el, data, successHtml, extraClass = '') {
  if (!el) return;
  el.className = `result-box ${data.success ? 'success' : 'error'} show ${extraClass}`;
  if (data.success) {
    el.innerHTML = `<div class="result-title">✅ Success</div><div class="result-details">${successHtml}</div>`;
  } else {
    el.innerHTML = `<div class="result-title">❌ Error</div><div class="result-details">${data.error || 'Something went wrong. Please try again.'}</div>`;
  }
  el.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

// ─── File Manager ─────────────────────────────────────────────
function initFileManager() {
  let selectedFiles = [];

  initDropZone('fm-drop-zone', 'fm-file-input', async (files) => {
    const formData = new FormData();
    files.forEach(f => formData.append('files', f));
    const progressEl = document.getElementById('fm-progress');
    const data = await uploadFiles('/api/file-manager/upload', formData, progressEl, null);
    if (data.success) {
      showToast(`✅ ${data.count} file(s) uploaded successfully`, 'success');
      refreshFileList();
    } else {
      showToast(data.error || 'Upload failed', 'error');
    }
  });

  const downloadAllBtn = document.getElementById('fm-download-all');
  if (downloadAllBtn) {
    downloadAllBtn.addEventListener('click', () => {
      window.location.href = '/api/file-manager/download-all';
    });
  }

  refreshFileList();
}

async function refreshFileList() {
  const listEl = document.getElementById('fm-file-list');
  if (!listEl) return;

  const res = await fetch('/api/file-manager/list');
  const data = await res.json();
  const files = data.files || [];

  if (files.length === 0) {
    listEl.innerHTML = `
      <div class="empty-state">
        <div class="empty-icon">📂</div>
        <h3>No files yet</h3>
        <p>Upload files using the drop zone above</p>
      </div>`;
    document.getElementById('fm-count').textContent = '0 files';
    return;
  }

  document.getElementById('fm-count').textContent = `${files.length} file${files.length > 1 ? 's' : ''}`;
  listEl.innerHTML = files.map(f => `
    <div class="file-item">
      <div class="file-ext-badge ext-${f.type.toLowerCase().replace('/', '')}">${f.ext.toUpperCase()}</div>
      <div class="file-info">
        <div class="file-name">${f.name}</div>
        <div class="file-meta">${f.size_str} &bull; ${f.modified}</div>
      </div>
      <div class="file-actions">
        <a href="/api/file-manager/download/${encodeURIComponent(f.name)}" class="btn btn-secondary btn-sm">⬇ Download</a>
        <button class="btn btn-danger btn-sm" onclick="deleteFile('${f.name}')">🗑 Delete</button>
      </div>
    </div>`).join('');
}

async function deleteFile(filename) {
  if (!confirm(`Delete "${filename}"?`)) return;
  const res = await fetch(`/api/file-manager/delete/${encodeURIComponent(filename)}`, { method: 'DELETE' });
  const data = await res.json();
  if (data.success) {
    showToast('File deleted', 'success');
    refreshFileList();
  } else {
    showToast(data.error || 'Delete failed', 'error');
  }
}

// ─── Module-level state ───────────────────────────────────────
let mergeFiles = [];

// ─── PDF Tools ────────────────────────────────────────────────
function initPDFTools() {
  // Compress
  initDropZone('compress-zone', 'compress-input', (files) => {
    if (files.length > 0) {
      document.getElementById('compress-filename').textContent = files[0].name;
      document.getElementById('compress-selected').style.display = 'block';
      document.getElementById('compress-btn').disabled = false;
      document.getElementById('compress-btn').dataset.file = 'ready';
    }
  });

  document.getElementById('compress-btn')?.addEventListener('click', async () => {
    const input = document.getElementById('compress-input');
    if (!input?.files[0]) { showToast('Please select a PDF file', 'warning'); return; }

    const formData = new FormData();
    formData.append('file', input.files[0]);
    formData.append('quality', document.getElementById('compress-quality')?.value || 'medium');

    const btn = document.getElementById('compress-btn');
    btn.disabled = true; btn.innerHTML = '<span class="spinner"></span> Compressing...';

    const data = await uploadFiles('/api/pdf/compress', formData, document.getElementById('compress-progress'), null);
    btn.disabled = false; btn.innerHTML = '🗜️ Compress PDF';

    const resultEl = document.getElementById('compress-result');
    if (data.success) {
      showResult(resultEl, data, `
        ${data.message}<br>
        <strong>Original:</strong> ${data.original_size_str} &rarr;
        <strong>Compressed:</strong> ${data.compressed_size_str}
        <span class="badge badge-success" style="margin-left:8px">-${data.reduction}%</span><br><br>
        <a href="/api/pdf/download/${data.filename}" class="btn btn-primary btn-sm">⬇ Download Compressed PDF</a>
      `);
      showToast(`PDF compressed! Size reduced by ${data.reduction}%`, 'success');
    } else {
      showResult(resultEl, data, '');
      showToast(data.error, 'error');
    }
  });

  // Merge
  mergeFiles = [];
  initDropZone('merge-zone', 'merge-input', (files) => {
    mergeFiles = mergeFiles.concat(files.filter(f => f.name.toLowerCase().endsWith('.pdf')));
    renderMergeList();
  });

  document.getElementById('merge-btn')?.addEventListener('click', async () => {
    if (mergeFiles.length < 2) { showToast('Please upload at least 2 PDF files', 'warning'); return; }

    const formData = new FormData();
    mergeFiles.forEach(f => formData.append('files', f));

    const btn = document.getElementById('merge-btn');
    btn.disabled = true; btn.innerHTML = '<span class="spinner"></span> Merging...';

    const data = await uploadFiles('/api/pdf/merge', formData, document.getElementById('merge-progress'), null);
    btn.disabled = false; btn.innerHTML = '🔗 Merge PDFs';

    const resultEl = document.getElementById('merge-result');
    if (data.success) {
      showResult(resultEl, data, `
        ${data.message}<br>
        <strong>Files merged:</strong> ${data.files_merged} &bull;
        <strong>Total pages:</strong> ${data.total_pages} &bull;
        <strong>Size:</strong> ${data.size_str}<br><br>
        <a href="/api/pdf/download/${data.filename}" class="btn btn-primary btn-sm">⬇ Download Merged PDF</a>
      `);
      showToast('PDFs merged successfully!', 'success');
    } else {
      showResult(resultEl, data, '');
      showToast(data.error, 'error');
    }
  });
}

function renderMergeList() {
  const listEl = document.getElementById('merge-file-list');
  if (!listEl) return;
  if (!mergeFiles || mergeFiles.length === 0) {
    listEl.innerHTML = '';
    document.getElementById('merge-btn').disabled = true;
    return;
  }
  document.getElementById('merge-btn').disabled = mergeFiles.length < 2;
  listEl.innerHTML = mergeFiles.map((f, i) => `
    <div class="file-item">
      <div class="file-ext-badge ext-pdf">PDF</div>
      <div class="file-info">
        <div class="file-name">${f.name}</div>
        <div class="file-meta">${formatBytes(f.size)}</div>
      </div>
      <div class="file-actions">
        <button class="btn btn-danger btn-sm" onclick="removeMergeFile(${i})">✕</button>
      </div>
    </div>`).join('');
}

function removeMergeFile(i) {
  mergeFiles.splice(i, 1);
  renderMergeList();
}


// ─── Excel Tools ──────────────────────────────────────────────
function initExcelTools() {
  let excelFiles = [];

  initDropZone('excel-zone', 'excel-input', (files) => {
    excelFiles = excelFiles.concat(files.filter(f => /\.(xlsx|xls|csv)$/i.test(f.name)));
    renderExcelList();
  });

  document.getElementById('excel-consolidate-btn')?.addEventListener('click', async () => {
    if (excelFiles.length === 0) { showToast('Please upload Excel or CSV files', 'warning'); return; }

    const formData = new FormData();
    excelFiles.forEach(f => formData.append('files', f));
    formData.append('mode', document.getElementById('consolidate-mode')?.value || 'append');

    const btn = document.getElementById('excel-consolidate-btn');
    btn.disabled = true; btn.innerHTML = '<span class="spinner"></span> Consolidating...';

    const data = await uploadFiles('/api/excel/consolidate', formData, document.getElementById('excel-progress'), null);
    btn.disabled = false; btn.innerHTML = '📊 Consolidate Files';

    const resultEl = document.getElementById('excel-result');
    if (data.success) {
      showResult(resultEl, data, `
        ${data.message}<br>
        <strong>Files:</strong> ${data.files_merged} &bull;
        <strong>Rows:</strong> ${data.total_rows} &bull;
        <strong>Size:</strong> ${data.size_str}<br><br>
        <a href="/api/excel/download/${data.filename}" class="btn btn-primary btn-sm">⬇ Download Excel</a>
      `);
      showToast('Files consolidated successfully!', 'success');
    } else {
      showResult(resultEl, data, '');
      showToast(data.error, 'error');
    }
  });

  function renderExcelList() {
    const listEl = document.getElementById('excel-file-list');
    if (!listEl) return;
    if (excelFiles.length === 0) { listEl.innerHTML = ''; return; }
    listEl.innerHTML = excelFiles.map((f, i) => `
      <div class="file-item">
        <div class="file-ext-badge ext-excel">${f.name.split('.').pop().toUpperCase()}</div>
        <div class="file-info">
          <div class="file-name">${f.name}</div>
          <div class="file-meta">${formatBytes(f.size)}</div>
        </div>
        <div class="file-actions">
          <button class="btn btn-danger btn-sm" onclick="removeExcelFile(${i})">✕</button>
        </div>
      </div>`).join('');
    document.getElementById('excel-consolidate-btn').disabled = false;
  }

  window.removeExcelFile = function(i) {
    excelFiles.splice(i, 1);
    renderExcelList();
  };
}

// ─── GST Tools ────────────────────────────────────────────────
function initGSTTools() {
  // GSTR-1
  document.getElementById('gstr1-btn')?.addEventListener('click', async () => {
    const input = document.getElementById('gstr1-input');
    if (!input?.files[0]) { showToast('Please select a file', 'warning'); return; }
    const formData = new FormData();
    formData.append('file', input.files[0]);
    const btn = document.getElementById('gstr1-btn');
    btn.disabled = true; btn.innerHTML = '<span class="spinner"></span> Processing...';
    const data = await uploadFiles('/api/gst/gstr1', formData, null, null);
    btn.disabled = false; btn.innerHTML = '📋 Process GSTR-1';
    const resultEl = document.getElementById('gstr1-result');
    if (data.success) {
      const totals = Object.entries(data.totals || {}).slice(0,4).map(([k,v]) => `<strong>${k}:</strong> ₹${v.toLocaleString('en-IN')}`).join(' &bull; ');
      showResult(resultEl, data, `${data.message}<br>${totals}<br><br><a href="/api/gst/download/${data.filename}" class="btn btn-primary btn-sm">⬇ Download Report</a>`);
      showToast('GSTR-1 processed!', 'success');
    } else { showResult(resultEl, data, ''); showToast(data.error, 'error'); }
  });

  // GSTR-2B Reco
  document.getElementById('gstr2b-btn')?.addEventListener('click', async () => {
    const pr = document.getElementById('gstr2b-pr-input');
    const b2b = document.getElementById('gstr2b-2b-input');
    if (!pr?.files[0] || !b2b?.files[0]) { showToast('Please upload both files', 'warning'); return; }
    const formData = new FormData();
    formData.append('purchase_file', pr.files[0]);
    formData.append('gstr2b_file', b2b.files[0]);
    const btn = document.getElementById('gstr2b-btn');
    btn.disabled = true; btn.innerHTML = '<span class="spinner"></span> Reconciling...';
    const data = await uploadFiles('/api/gst/gstr2b-reco', formData, null, null);
    btn.disabled = false; btn.innerHTML = '🔄 Reconcile';
    const resultEl = document.getElementById('gstr2b-result');
    if (data.success) {
      showResult(resultEl, data, `
        ${data.message}<br>
        <div class="reco-summary" style="margin-top:12px">
          <div class="reco-card reco-matched"><div class="reco-num">${data.matched}</div><div class="reco-lbl">Matched</div></div>
          <div class="reco-card reco-only-pr"><div class="reco-num">${data.only_in_pr}</div><div class="reco-lbl">Only in PR</div></div>
          <div class="reco-card reco-only-2b"><div class="reco-num">${data.only_in_2b}</div><div class="reco-lbl">Only in 2B</div></div>
          <div class="reco-card reco-total"><div class="reco-num">${data.total_pr}</div><div class="reco-lbl">Total PR</div></div>
        </div><br>
        <a href="/api/gst/download/${data.filename}" class="btn btn-primary btn-sm">⬇ Download Reconciliation</a>
      `);
      showToast('Reconciliation complete!', 'success');
    } else { showResult(resultEl, data, ''); showToast(data.error, 'error'); }
  });

  // GSTR-3B
  document.getElementById('gstr3b-btn')?.addEventListener('click', async () => {
    const input = document.getElementById('gstr3b-input');
    if (!input?.files[0]) { showToast('Please select a file', 'warning'); return; }
    const formData = new FormData();
    formData.append('file', input.files[0]);
    formData.append('period', document.getElementById('gstr3b-period')?.value || '');
    const btn = document.getElementById('gstr3b-btn');
    btn.disabled = true; btn.innerHTML = '<span class="spinner"></span> Processing...';
    const data = await uploadFiles('/api/gst/gstr3b', formData, null, null);
    btn.disabled = false; btn.innerHTML = '📋 Process GSTR-3B';
    const resultEl = document.getElementById('gstr3b-result');
    if (data.success) {
      showResult(resultEl, data, `${data.message}<br><a href="/api/gst/download/${data.filename}" class="btn btn-primary btn-sm" style="margin-top:12px">⬇ Download Report</a>`);
      showToast('GSTR-3B processed!', 'success');
    } else { showResult(resultEl, data, ''); showToast(data.error, 'error'); }
  });

  // GSTR-9
  document.getElementById('gstr9-btn')?.addEventListener('click', async () => {
    const input = document.getElementById('gstr9-input');
    if (!input?.files[0]) { showToast('Please select a file', 'warning'); return; }
    const formData = new FormData();
    formData.append('file', input.files[0]);
    formData.append('fy', document.getElementById('gstr9-fy')?.value || '');
    const btn = document.getElementById('gstr9-btn');
    btn.disabled = true; btn.innerHTML = '<span class="spinner"></span> Processing...';
    const data = await uploadFiles('/api/gst/gstr9', formData, null, null);
    btn.disabled = false; btn.innerHTML = '📋 Process GSTR-9';
    const resultEl = document.getElementById('gstr9-result');
    if (data.success) {
      showResult(resultEl, data, `${data.message}<br><a href="/api/gst/download/${data.filename}" class="btn btn-primary btn-sm" style="margin-top:12px">⬇ Download Report</a>`);
      showToast('GSTR-9 processed!', 'success');
    } else { showResult(resultEl, data, ''); showToast(data.error, 'error'); }
  });

  // PR vs 2B
  document.getElementById('pr2b-btn')?.addEventListener('click', async () => {
    const pr = document.getElementById('pr2b-pr-input');
    const b2b = document.getElementById('pr2b-2b-input');
    if (!pr?.files[0] || !b2b?.files[0]) { showToast('Please upload both files', 'warning'); return; }
    const formData = new FormData();
    formData.append('pr_file', pr.files[0]);
    formData.append('b2b_file', b2b.files[0]);
    const btn = document.getElementById('pr2b-btn');
    btn.disabled = true; btn.innerHTML = '<span class="spinner"></span> Reconciling...';
    const data = await uploadFiles('/api/gst/pr-2b-reco', formData, null, null);
    btn.disabled = false; btn.innerHTML = '🔄 Run Reconciliation';
    const resultEl = document.getElementById('pr2b-result');
    if (data.success) {
      showResult(resultEl, data, `
        ${data.message}<br>
        <div class="reco-summary" style="margin-top:12px">
          <div class="reco-card reco-matched"><div class="reco-num">${data.matched}</div><div class="reco-lbl">Matched</div></div>
          <div class="reco-card reco-only-pr"><div class="reco-num">${data.only_in_pr}</div><div class="reco-lbl">Only in PR</div></div>
          <div class="reco-card reco-only-2b"><div class="reco-num">${data.only_in_2b}</div><div class="reco-lbl">Only in 2B</div></div>
        </div><br>
        <a href="/api/gst/download/${data.filename}" class="btn btn-primary btn-sm">⬇ Download 4-Sheet Report</a>
      `);
      showToast('PR vs 2B Reconciliation complete!', 'success');
    } else { showResult(resultEl, data, ''); showToast(data.error, 'error'); }
  });
}

// ─── Init ──────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  initSidebar();
  initTabs('tabs-container');

  const page = document.body.dataset.page;
  if (page === 'file-manager') initFileManager();
  if (page === 'pdf-tools') initPDFTools();
  if (page === 'excel-tools') initExcelTools();
  if (page === 'gst-tools') initGSTTools();
});
