/**
 * PEA Scrap Metal Return Calculator - Application Logic
 * Provincial Electricity Authority (กฟภ.)
 */

// Application State
const state = {
  masterData: [],
  metadata: {
    job_no: '',
    job_name: '',
    person_name: '',
    person_id: '',
    cost_center_resp: '',
    print_date: ''
  },
  items: [],
  loadedFileName: '',
  sectionFilter: 'dismantle_only',
  matchFilter: 'exact',
  qtySource: 'suggested',
  searchQuery: '',
  scrapPricePerKg: 12.0,
  isServerOnline: false,
  availableDesktopFiles: []
};

// DOM Elements
const elements = {
  navTabs: document.querySelectorAll('.nav-tab'),
  tabContents: document.querySelectorAll('.tab-content'),
  dropZone: document.getElementById('dropZone'),
  fileInput: document.getElementById('fileInput'),
  browseBtn: document.getElementById('browseBtn'),
  quickSampleLink: document.getElementById('quickSampleLink'),
  quickLoadDesktopBtn: document.getElementById('quickLoadDesktopBtn'),
  desktopFileSelect: document.getElementById('desktopFileSelect'),
  desktopSamplePicker: document.getElementById('desktopSamplePicker'),
  offlineAlertBanner: document.getElementById('offlineAlertBanner'),
  retryServerBtn: document.getElementById('retryServerBtn'),
  sampleLinksContainer: document.getElementById('sampleLinksContainer'),
  loadingOverlay: document.getElementById('loadingOverlay'),
  loadingText: document.getElementById('loadingText'),
  projectBanner: document.getElementById('projectBanner'),
  loadedFileName: document.getElementById('loadedFileName'),
  metaJobNo: document.getElementById('metaJobNo'),
  metaJobName: document.getElementById('metaJobName'),
  metaPersonName: document.getElementById('metaPersonName'),
  metaPersonId: document.getElementById('metaPersonId'),
  metaCostCenter: document.getElementById('metaCostCenter'),
  metaPrintDate: document.getElementById('metaPrintDate'),
  totalKgDisplay: document.getElementById('totalKgDisplay'),
  totalTonsDisplay: document.getElementById('totalTonsDisplay'),
  totalValueDisplay: document.getElementById('totalValueDisplay'),
  totalItemsDisplay: document.getElementById('totalItemsDisplay'),
  matchStatsDisplay: document.getElementById('matchStatsDisplay'),
  matchRatioDisplay: document.getElementById('matchRatioDisplay'),
  scrapPriceInput: document.getElementById('scrapPriceInput'),
  sectionFilter: document.getElementById('sectionFilter'),
  matchFilter: document.getElementById('matchFilter'),
  qtySourceSelect: document.getElementById('qtySourceSelect'),
  tableSearchInput: document.getElementById('tableSearchInput'),
  checkAllBtn: document.getElementById('checkAllBtn'),
  uncheckAllBtn: document.getElementById('uncheckAllBtn'),
  masterCheckbox: document.getElementById('masterCheckbox'),
  addManualItemBtn: document.getElementById('addManualItemBtn'),
  exportExcelBtn: document.getElementById('exportExcelBtn'),
  printReportBtn: document.getElementById('printReportBtn'),
  calcTable: document.getElementById('calcTable'),
  calcTableBody: document.getElementById('calcTableBody'),
  calcTableFoot: document.getElementById('calcTableFoot'),
  footTotalQty: document.getElementById('footTotalQty'),
  footTotalWeight: document.getElementById('footTotalWeight'),
  footTotalTons: document.getElementById('footTotalTons'),
  masterTableBody: document.getElementById('masterTableBody'),
  masterTotalItems: document.getElementById('masterTotalItems'),
  masterCountBadge: document.getElementById('masterCountBadge'),
  masterSearchInput: document.getElementById('masterSearchInput'),
  addMasterItemBtn: document.getElementById('addMasterItemBtn'),
  resetMasterBtn: document.getElementById('resetMasterBtn'),
  printTableBody: document.getElementById('printTableBody'),
  printGrandTotalKg: document.getElementById('printGrandTotalKg'),
  printGrandTotalTons: document.getElementById('printGrandTotalTons'),
  printJobNo: document.getElementById('printJobNo'),
  printJobName: document.getElementById('printJobName'),
  printPersonName: document.getElementById('printPersonName'),
  printPersonId: document.getElementById('printPersonId'),
  printCostCenter: document.getElementById('printCostCenter'),
  printDateVal: document.getElementById('printDateVal'),
  printThaiWordSum: document.getElementById('printThaiWordSum'),
  signPersonName: document.getElementById('signPersonName'),
  printReportSheet: document.getElementById('printReportSheet'),
  toggleFitOnePage: document.getElementById('toggleFitOnePage'),
  printScaleSelect: document.getElementById('printScaleSelect'),
  manualItemModal: document.getElementById('manualItemModal'),
  closeManualModalBtn: document.getElementById('closeManualModalBtn'),
  cancelManualModalBtn: document.getElementById('cancelManualModalBtn'),
  saveManualItemBtn: document.getElementById('saveManualItemBtn'),
  manualCodeInput: document.getElementById('manualCodeInput'),
  manualDescInput: document.getElementById('manualDescInput'),
  manualQtyInput: document.getElementById('manualQtyInput'),
  manualUnitInput: document.getElementById('manualUnitInput'),
  manualWeightInput: document.getElementById('manualWeightInput'),
  manualMasterSelect: document.getElementById('manualMasterSelect'),
  masterItemModal: document.getElementById('masterItemModal'),
  closeMasterModalBtn: document.getElementById('closeMasterModalBtn'),
  cancelMasterModalBtn: document.getElementById('cancelMasterModalBtn'),
  saveMasterItemBtn: document.getElementById('saveMasterItemBtn'),
  masterItemCode: document.getElementById('masterItemCode'),
  masterItemName: document.getElementById('masterItemName'),
  masterItemWeight: document.getElementById('masterItemWeight'),
  masterItemUnit: document.getElementById('masterItemUnit'),
  masterModalTitle: document.getElementById('masterModalTitle'),
  masterEditId: document.getElementById('masterEditId'),
  toastContainer: document.getElementById('toastContainer')
};

// Initialize Application
async function init() {
  setupEventListeners();
  await loadMasterData();
  checkServerHealth();
}

// Load Master Scrap Weights
async function loadMasterData() {
  try {
    const res = await fetch('/api/master-data');
    if (res.ok) {
      state.masterData = await res.json();
    } else {
      // Fallback load local master_data.json
      const localRes = await fetch('master_data.json');
      state.masterData = await localRes.json();
    }
  } catch (err) {
    console.warn('Could not load from API, falling back to local master_data.json:', err);
    try {
      const localRes = await fetch('master_data.json');
      state.masterData = await localRes.json();
    } catch (e) {
      console.error('Failed to load master data:', e);
      state.masterData = [];
    }
  }

  // Update badges and dropdowns
  elements.masterCountBadge.textContent = state.masterData.length;
  elements.masterTotalItems.textContent = state.masterData.length;
  renderMasterTable();
  populateManualModalSelect();
}

// Check Server Health
async function checkServerHealth() {
  const statusBadge = document.getElementById('serverStatusBadge');
  const statusText = document.getElementById('serverStatusText');
  const statusDot = statusBadge ? statusBadge.querySelector('.status-dot') : null;
  const offlineBanner = elements.offlineAlertBanner;

  try {
    const res = await fetch('/api/status');
    if (res.ok) {
      const data = await res.json();
      state.isServerOnline = true;
      if (statusBadge) {
        statusBadge.classList.remove('offline');
        statusBadge.title = 'เซิร์ฟเวอร์พร้อมทำงาน';
      }
      if (statusDot) {
        statusDot.className = 'status-dot online';
      }
      if (statusText) {
        statusText.textContent = data.is_vercel ? 'ระบบพร้อมใช้งาน (Vercel Cloud)' : 'ระบบพร้อมใช้งาน';
      }
      if (offlineBanner) {
        offlineBanner.style.display = 'none';
      }

      if (data.available_files && data.available_files.length > 0) {
        state.availableDesktopFiles = data.available_files;
        renderDesktopFileSelector(data.available_files);
      }
      return true;
    }
  } catch (e) {
    console.warn('Server is offline:', e);
  }

  // If reached here, server is offline
  state.isServerOnline = false;
  const isLocalhost = window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1';
  const offlineAlertText = document.getElementById('offlineAlertText');

  if (offlineAlertText) {
    if (isLocalhost) {
      offlineAlertText.innerHTML = '<strong>เซิร์ฟเวอร์หลังบ้านออฟไลน์:</strong> ระบบประมวลผลจำเป็นต้องเปิดไฟล์ <code>start_app.bat</code> เพื่อเริ่มการทำงาน';
    } else {
      offlineAlertText.innerHTML = '<strong>กำลังเชื่อมต่อ Cloud Serverless:</strong> หากเปิดเว็บเป็นครั้งแรก Vercel กำลังเตรียมฟังก์ชันประมวลผล กรุณารอสักครู่แล้วกดปุ่ม <em>"ตรวจสอบการเชื่อมต่อ"</em>';
    }
  }

  if (statusBadge) {
    statusBadge.classList.add('offline');
    statusBadge.title = isLocalhost ? 'เซิร์ฟเวอร์ Local ยังไม่เปิด (start_app.bat)' : 'กำลังเชื่อมต่อ Cloud Serverless';
  }
  if (statusDot) {
    statusDot.className = 'status-dot offline';
  }
  if (statusText) {
    statusText.textContent = isLocalhost ? 'เซิร์ฟเวอร์ออฟไลน์' : 'กำลังเชื่อมต่อ Cloud...';
  }
  if (offlineBanner) {
    offlineBanner.style.display = 'block';
  }
  return false;
}

// Render Desktop 018 file selector
function renderDesktopFileSelector(files) {
  if (!elements.desktopFileSelect) return;

  elements.desktopFileSelect.innerHTML = '';
  files.forEach(f => {
    const opt = document.createElement('option');
    opt.value = f.filename;
    opt.textContent = `⚡ ${f.label || f.filename} (${f.size_kb} KB)`;
    elements.desktopFileSelect.appendChild(opt);
  });

  // Render quick sample links in dropzone
  if (elements.sampleLinksContainer) {
    elements.sampleLinksContainer.innerHTML = files.map(f => {
      const shortName = f.label ? f.label.replace(/\(.*?\)/g, '').trim() : f.filename;
      return `<a href="javascript:void(0)" class="quick-file-pill" data-filename="${f.filename}" style="display: inline-block; margin: 2px 6px; padding: 2px 10px; background: rgba(94, 19, 110, 0.08); border: 1px solid rgba(94, 19, 110, 0.2); border-radius: 20px; font-size: 12px; text-decoration: none; color: #5E136E; font-weight: 600;">⚡ โหลด ${shortName}</a>`;
    }).join('');

    elements.sampleLinksContainer.querySelectorAll('.quick-file-pill').forEach(pill => {
      pill.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        loadDesktopSample(pill.dataset.filename);
      });
    });
  }
}

// Quick Load Desktop Sample
const loadDesktopSample = async (filename) => {
  const targetName = filename || (elements.desktopFileSelect ? elements.desktopFileSelect.value : '') || '018.pdf';
  showLoading(`กำลังโหลดไฟล์ ${targetName} จาก Desktop...`);
  try {
    const res = await fetch(`/api/load-desktop-sample?file=${encodeURIComponent(targetName)}`);
    const data = await res.json();
    if (data.success) {
      processLoadedData(data);
      showToast(`โหลดข้อมูลจาก ${data.filename} สำเร็จเรียบร้อย!`, 'success');
    } else {
      showToast(data.error || 'ไม่สามารถโหลดไฟล์ได้', 'error', 7000);
    }
  } catch (err) {
    showToast(`ไม่สามารถเชื่อมต่อเซิร์ฟเวอร์: ${err.message} กรุณาเปิดไฟล์ start_app.bat`, 'error', 7000);
    checkServerHealth();
  } finally {
    hideLoading();
  }
};

// Event Listeners Setup
function setupEventListeners() {
  // Navigation Tabs
  elements.navTabs.forEach(tab => {
    tab.addEventListener('click', () => {
      elements.navTabs.forEach(t => t.classList.remove('active'));
      elements.tabContents.forEach(c => c.classList.remove('active'));
      tab.classList.add('active');
      const target = document.getElementById(tab.dataset.tab);
      if (target) target.classList.add('active');

      if (tab.dataset.tab === 'printTab') {
        renderPrintReport();
      }
    });
  });

  // Print Fit & Scale Listeners
  if (elements.toggleFitOnePage) {
    elements.toggleFitOnePage.addEventListener('change', () => {
      applyPrintFit();
    });
  }

  if (elements.printScaleSelect) {
    elements.printScaleSelect.addEventListener('change', () => {
      applyPrintFit();
    });
  }

  // Server retry button
  if (elements.retryServerBtn) {
    elements.retryServerBtn.addEventListener('click', async () => {
      const ok = await checkServerHealth();
      if (ok) {
        showToast('เชื่อมต่อเซิร์ฟเวอร์สำเร็จเรียบร้อย!', 'success');
      } else {
        showToast('ยังไม่สามารถเชื่อมต่อเซิร์ฟเวอร์ได้ กรุณาตรวจสอบ start_app.bat', 'error');
      }
    });
  }

  // Upload actions
  elements.browseBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    elements.fileInput.click();
  });
  elements.dropZone.addEventListener('click', () => elements.fileInput.click());

  elements.fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
      handleFileUpload(e.target.files[0]);
    }
  });

  // Drag & Drop
  ['dragenter', 'dragover'].forEach(name => {
    elements.dropZone.addEventListener(name, (e) => {
      e.preventDefault();
      e.stopPropagation();
      elements.dropZone.classList.add('drag-over');
    });
  });

  ['dragleave', 'drop'].forEach(name => {
    elements.dropZone.addEventListener(name, (e) => {
      e.preventDefault();
      e.stopPropagation();
      elements.dropZone.classList.remove('drag-over');
    });
  });

  elements.dropZone.addEventListener('drop', (e) => {
    if (e.dataTransfer.files.length > 0) {
      handleFileUpload(e.dataTransfer.files[0]);
    }
  });

  if (elements.quickSampleLink) {
    elements.quickSampleLink.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      loadDesktopSample();
    });
  }

  if (elements.quickLoadDesktopBtn) {
    elements.quickLoadDesktopBtn.addEventListener('click', () => {
      const selected = elements.desktopFileSelect ? elements.desktopFileSelect.value : '';
      loadDesktopSample(selected);
    });
  }

  if (elements.desktopFileSelect) {
    elements.desktopFileSelect.addEventListener('change', (e) => {
      if (e.target.value) {
        loadDesktopSample(e.target.value);
      }
    });
  }

  // Filters and Toolbar
  elements.sectionFilter.addEventListener('change', (e) => {
    state.sectionFilter = e.target.value;
    renderCalcTable();
  });

  elements.matchFilter.addEventListener('change', (e) => {
    state.matchFilter = e.target.value;
    renderCalcTable();
  });

  elements.tableSearchInput.addEventListener('input', (e) => {
    state.searchQuery = e.target.value.toLowerCase().trim();
    renderCalcTable();
  });

  elements.qtySourceSelect.addEventListener('change', (e) => {
    state.qtySource = e.target.value;
    applyQtySourceToAll();
  });

  elements.scrapPriceInput.addEventListener('input', (e) => {
    state.scrapPricePerKg = parseFloat(e.target.value) || 0;
    updateKPICards();
  });

  elements.checkAllBtn.addEventListener('click', () => toggleAllVisible(true));
  elements.uncheckAllBtn.addEventListener('click', () => toggleAllVisible(false));
  elements.masterCheckbox.addEventListener('change', (e) => toggleAllVisible(e.target.checked));

  // Export and Print
  elements.exportExcelBtn.addEventListener('click', handleExportExcel);
  elements.printReportBtn.addEventListener('click', () => {
    const printNavTab = document.querySelector('.nav-tab[data-tab="printTab"]');
    if (printNavTab) printNavTab.click();
  });

  // Manual Item Modal
  elements.addManualItemBtn.addEventListener('click', () => {
    elements.manualCodeInput.value = '';
    elements.manualDescInput.value = '';
    elements.manualQtyInput.value = '1';
    elements.manualUnitInput.value = 'EA';
    elements.manualWeightInput.value = '1.0';
    elements.manualMasterSelect.value = '';
    elements.manualItemModal.style.display = 'flex';
  });

  elements.closeManualModalBtn.addEventListener('click', () => elements.manualItemModal.style.display = 'none');
  elements.cancelManualModalBtn.addEventListener('click', () => elements.manualItemModal.style.display = 'none');
  elements.saveManualItemBtn.addEventListener('click', handleSaveManualItem);

  elements.manualMasterSelect.addEventListener('change', (e) => {
    const selectedMaster = state.masterData.find(m => m.id == e.target.value);
    if (selectedMaster) {
      elements.manualCodeInput.value = selectedMaster.formatted_code || selectedMaster.code;
      elements.manualDescInput.value = selectedMaster.name;
      elements.manualWeightInput.value = selectedMaster.weight_per_unit;
      elements.manualUnitInput.value = selectedMaster.unit;
    }
  });

  // Master Data Tab actions
  elements.masterSearchInput.addEventListener('input', renderMasterTable);
  elements.addMasterItemBtn.addEventListener('click', () => {
    elements.masterModalTitle.textContent = 'เพิ่มอุปกรณ์ใหม่ในฐานข้อมูล';
    elements.masterEditId.value = '';
    elements.masterItemCode.value = '';
    elements.masterItemName.value = '';
    elements.masterItemWeight.value = '1.0';
    elements.masterItemUnit.value = 'กก.';
    elements.masterItemModal.style.display = 'flex';
  });

  elements.closeMasterModalBtn.addEventListener('click', () => elements.masterItemModal.style.display = 'none');
  elements.cancelMasterModalBtn.addEventListener('click', () => elements.masterItemModal.style.display = 'none');
  elements.saveMasterItemBtn.addEventListener('click', handleSaveMasterItem);

  elements.resetMasterBtn.addEventListener('click', async () => {
    if (confirm('คุณต้องการคืนค่าฐานข้อมูลน้ำหนักอุปกรณ์เป็นค่าเริ่มต้นจาก Desktop Excel ใช่หรือไม่?')) {
      await loadMasterData();
      showToast('คืนค่าฐานข้อมูลน้ำหนักเริ่มต้นเรียบร้อย', 'success');
    }
  });
}

// File Upload Handler
async function handleFileUpload(file) {
  if (!file) return;
  const ext = file.name.split('.').pop().toLowerCase();
  if (ext !== 'pdf' && ext !== 'xlsx' && ext !== 'xls') {
    showToast('กรุณาอัปโหลดไฟล์นามสกุล .pdf หรือ .xlsx จากระบบ SAP', 'error');
    return;
  }

  showLoading(`กำลังวิเคราะห์ไฟล์ ${file.name}...`);
  const formData = new FormData();
  formData.append('file', file);

  try {
    const res = await fetch('/api/parse-pdf', {
      method: 'POST',
      body: formData
    });
    const data = await res.json();
    if (data.success) {
      processLoadedData(data);
      showToast(`วิเคราะห์ไฟล์ ${file.name} สำเร็จ! พบรายการพัสดุ ${data.items.length} รายการ`, 'success');
    } else {
      showToast(data.error || 'เกิดข้อผิดพลาดในการอ่านไฟล์', 'error', 8000);
    }
  } catch (err) {
    showToast(`ไม่สามารถส่งไฟล์ไปยังเซิร์ฟเวอร์: ${err.message} กรุณาตรวจสอบว่าได้เปิด start_app.bat แล้วหรือยัง`, 'error', 8000);
    checkServerHealth();
  } finally {
    hideLoading();
  }
}

// Process Loaded Data from SAP Report
function processLoadedData(data) {
  state.loadedFileName = data.filename || '018.pdf';
  state.metadata = data.metadata || {};
  state.items = (data.items || []).map((it, idx) => ({
    ...it,
    id: idx + 1
  }));

  // Update UI Metadata
  elements.projectBanner.style.display = 'block';
  elements.loadedFileName.textContent = state.loadedFileName;
  elements.metaJobNo.textContent = state.metadata.job_no || '-';
  elements.metaJobName.textContent = state.metadata.job_name || '-';
  elements.metaPersonName.textContent = state.metadata.person_name || '-';
  elements.metaPersonId.textContent = state.metadata.person_id || '-';
  elements.metaCostCenter.textContent = state.metadata.cost_center_resp || '-';
  elements.metaPrintDate.textContent = state.metadata.print_date || '-';

  // Populate Section Filter dynamically with all sections found
  populateSectionFilter();

  // Render Table & KPI Cards
  renderCalcTable();
  updateKPICards();
}

// Populate Section Filter
function populateSectionFilter() {
  const sections = Array.from(new Set(state.items.map(it => it.section).filter(Boolean)));
  const hasDismantle = state.items.some(it => it.is_dismantle);
  const dismantleCount = state.items.filter(it => it.is_dismantle).length;

  // Auto-switch filter: if no dismantle items found, show all so table is never empty!
  if (dismantleCount === 0 && state.items.length > 0) {
    state.sectionFilter = 'all';
  } else {
    state.sectionFilter = 'dismantle_only';
  }

  elements.sectionFilter.innerHTML = `
    <option value="dismantle_only" ${state.sectionFilter === 'dismantle_only' ? 'selected' : ''}>⚡ เฉพาะงานรื้อถอน (-R-E) (${dismantleCount} รายการ)</option>
    <option value="all" ${state.sectionFilter === 'all' ? 'selected' : ''}>ทั้งหมด ทุกแผนกงาน (${state.items.length} รายการ)</option>
  `;

  sections.forEach(sec => {
    const count = state.items.filter(it => it.section === sec).length;
    const isDismantle = sec.includes('-R-E') || sec.includes('รื้อถอน') || sec.includes('ถอน');
    const opt = document.createElement('option');
    opt.value = sec;
    opt.textContent = `${isDismantle ? '🔨 ' : '🏗️ '}${sec} (${count})`;
    elements.sectionFilter.appendChild(opt);
  });
}

// Apply selected Quantity column across all items
function applyQtySourceToAll() {
  state.items.forEach(it => {
    if (state.qtySource === 'suggested') {
      it.calc_qty = it.suggested_qty;
    } else if (state.qtySource === 'estimate') {
      it.calc_qty = it.qty_estimate;
    } else if (state.qtySource === 'good_return') {
      it.calc_qty = it.qty_good_return;
    } else if (state.qtySource === 'damaged_return') {
      it.calc_qty = it.qty_damaged_return;
    } else if (state.qtySource === 'installed') {
      it.calc_qty = it.qty_installed;
    }
    it.total_weight = Math.round((it.calc_qty * it.weight_per_unit) * 1000) / 1000;
  });
  renderCalcTable();
}

// Get filtered items based on active filters
function getFilteredItems() {
  return state.items.filter(it => {
    // 1. Section filter
    if (state.sectionFilter === 'dismantle_only') {
      if (!it.is_dismantle) return false;
    } else if (state.sectionFilter !== 'all') {
      if (it.section !== state.sectionFilter) return false;
    }

    // 2. Match filter
    if (state.matchFilter !== 'all') {
      if (it.match_type !== state.matchFilter) return false;
    }

    // 3. Search query
    if (state.searchQuery) {
      const q = state.searchQuery;
      const matchSearch =
        (it.code && it.code.toLowerCase().includes(q)) ||
        (it.desc && it.desc.toLowerCase().includes(q)) ||
        (it.master_name && it.master_name.toLowerCase().includes(q)) ||
        (it.section && it.section.toLowerCase().includes(q));
      if (!matchSearch) return false;
    }

    return true;
  });
}

// Render Main Calculation Table
function renderCalcTable() {
  const filtered = getFilteredItems();
  elements.calcTableBody.innerHTML = '';

  if (filtered.length === 0) {
    elements.calcTableBody.innerHTML = `
      <tr>
        <td colspan="12" class="empty-row">
          <div class="empty-state">
            <div class="empty-icon">🔍</div>
            <div class="empty-text">ไม่พบรายการพัสดุตามเงื่อนไขตัวกรอง</div>
            <div class="empty-sub">ลองปรับเปลี่ยนตัวกรองแผนก หรือพิมพ์คำค้นหาใหม่</div>
          </div>
        </td>
      </tr>
    `;
    elements.calcTableFoot.style.display = 'none';
    updateKPICards();
    return;
  }

  elements.calcTableFoot.style.display = 'table-footer-group';

  filtered.forEach((it, index) => {
    const tr = document.createElement('tr');
    if (it.selected) tr.classList.add('row-selected');

    // Match badge
    let badgeHtml = '';
    if (it.match_type === 'exact') {
      badgeHtml = '<span class="badge badge-exact">🟢 ตรงกัน</span>';
    } else if (it.match_type === 'suggested') {
      badgeHtml = '<span class="badge badge-suggested">🟡 แนะนำ</span>';
    } else {
      badgeHtml = '<span class="badge badge-none">⚪ นอกฐาน</span>';
    }

    // Section tag
    const isRe = it.is_dismantle;
    const secTag = `<span class="section-tag ${isRe ? 'dismantle' : 'construction'}">${it.section ? it.section.split(' ')[1] || it.section : '-'}</span>`;

    tr.innerHTML = `
      <td class="text-center">
        <input type="checkbox" class="row-checkbox" data-id="${it.id}" ${it.selected ? 'checked' : ''}>
      </td>
      <td class="text-center">${index + 1}</td>
      <td><span class="cell-code">${it.code || '-'}</span></td>
      <td><div class="cell-desc" title="${it.desc}">${it.desc || '-'}</div></td>
      <td>
        <select class="cell-master-select" data-id="${it.id}">
          <option value="">-- เลือกอุปกรณ์อ้างอิง --</option>
          ${state.masterData.map(m => `
            <option value="${m.id}" ${m.name === it.master_name ? 'selected' : ''}>
              ${m.name} (${m.weight_per_unit} กก.)
            </option>
          `).join('')}
        </select>
      </td>
      <td>${secTag}</td>
      <td class="text-center">${it.unit || 'EA'}</td>
      <td class="text-right">
        <input type="number" class="cell-qty-input" data-id="${it.id}" value="${it.calc_qty}" min="0" step="any">
      </td>
      <td class="text-right">
        <input type="number" class="cell-wt-input" data-id="${it.id}" value="${it.weight_per_unit}" min="0" step="any">
      </td>
      <td class="text-right cell-total-weight" id="itemTotal_${it.id}">
        ${(it.total_weight || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
      </td>
      <td class="text-center">${badgeHtml}</td>
      <td class="text-center">
        <button class="btn-delete-row" data-id="${it.id}" title="ลบรายการนี้">&times;</button>
      </td>
    `;

    elements.calcTableBody.appendChild(tr);
  });

  // Attach Table Row Event Listeners
  attachTableRowListeners();
  updateKPICards();
}

// Table Row Event Listeners
function attachTableRowListeners() {
  // Row checkboxes
  document.querySelectorAll('.row-checkbox').forEach(cb => {
    cb.addEventListener('change', (e) => {
      const id = parseInt(e.target.dataset.id);
      const item = state.items.find(it => it.id === id);
      if (item) {
        item.selected = e.target.checked;
        const tr = e.target.closest('tr');
        if (tr) tr.classList.toggle('row-selected', item.selected);
        updateKPICards();
      }
    });
  });

  // Quantity input
  document.querySelectorAll('.cell-qty-input').forEach(input => {
    input.addEventListener('input', (e) => {
      const id = parseInt(e.target.dataset.id);
      const item = state.items.find(it => it.id === id);
      if (item) {
        item.calc_qty = parseFloat(e.target.value) || 0;
        item.total_weight = Math.round((item.calc_qty * item.weight_per_unit) * 1000) / 1000;
        const totalCell = document.getElementById(`itemTotal_${id}`);
        if (totalCell) {
          totalCell.textContent = item.total_weight.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        }
        updateKPICards();
      }
    });
  });

  // Weight per unit input
  document.querySelectorAll('.cell-wt-input').forEach(input => {
    input.addEventListener('input', (e) => {
      const id = parseInt(e.target.dataset.id);
      const item = state.items.find(it => it.id === id);
      if (item) {
        item.weight_per_unit = parseFloat(e.target.value) || 0;
        item.total_weight = Math.round((item.calc_qty * item.weight_per_unit) * 1000) / 1000;
        const totalCell = document.getElementById(`itemTotal_${id}`);
        if (totalCell) {
          totalCell.textContent = item.total_weight.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        }
        updateKPICards();
      }
    });
  });

  // Master dropdown change
  document.querySelectorAll('.cell-master-select').forEach(select => {
    select.addEventListener('change', (e) => {
      const id = parseInt(e.target.dataset.id);
      const item = state.items.find(it => it.id === id);
      const masterId = parseInt(e.target.value);
      const master = state.masterData.find(m => m.id === masterId);
      if (item && master) {
        item.master_name = master.name;
        item.weight_per_unit = master.weight_per_unit;
        item.match_type = 'suggested';
        item.total_weight = Math.round((item.calc_qty * item.weight_per_unit) * 1000) / 1000;
        renderCalcTable();
      }
    });
  });

  // Delete row
  document.querySelectorAll('.btn-delete-row').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const id = parseInt(e.target.dataset.id);
      state.items = state.items.filter(it => it.id !== id);
      renderCalcTable();
    });
  });
}

// Toggle visible items
function toggleAllVisible(checked) {
  const filtered = getFilteredItems();
  filtered.forEach(it => {
    it.selected = checked;
  });
  renderCalcTable();
}

// Update KPI Summary Cards & Table Footer
function updateKPICards() {
  const selectedItems = state.items.filter(it => it.selected);
  const totalKg = selectedItems.reduce((acc, it) => acc + (it.total_weight || 0), 0);
  const totalTons = totalKg / 1000.0;
  const totalValue = totalKg * state.scrapPricePerKg;
  const totalSelectedCount = selectedItems.length;

  elements.totalKgDisplay.innerHTML = `${totalKg.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} <span class="kpi-unit">กก.</span>`;
  elements.totalTonsDisplay.innerHTML = `${totalTons.toLocaleString('en-US', { minimumFractionDigits: 3, maximumFractionDigits: 3 })} <span class="kpi-unit">ตัน</span>`;
  elements.totalValueDisplay.innerHTML = `${totalValue.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} <span class="kpi-unit">บาท</span>`;
  elements.totalItemsDisplay.textContent = `${totalSelectedCount} รายการที่เลือกคำนวณ (จาก ${state.items.length})`;

  // Match statistics
  const matchedCount = state.items.filter(it => it.match_type === 'exact' || it.match_type === 'suggested').length;
  elements.matchStatsDisplay.textContent = `${matchedCount} / ${state.items.length}`;
  const pct = state.items.length ? Math.round((matchedCount / state.items.length) * 100) : 0;
  elements.matchRatioDisplay.textContent = `ตรงกับฐานข้อมูล ${pct}%`;

  // Table Footer
  const totalQty = selectedItems.reduce((acc, it) => acc + (it.calc_qty || 0), 0);
  elements.footTotalQty.textContent = totalQty.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  elements.footTotalWeight.textContent = `${totalKg.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} กก.`;
  elements.footTotalTons.textContent = `${totalTons.toLocaleString('en-US', { minimumFractionDigits: 3, maximumFractionDigits: 3 })} ตัน`;
}

// Render Master Data Table
function renderMasterTable() {
  const query = (elements.masterSearchInput.value || '').toLowerCase().trim();
  const list = state.masterData.filter(m => {
    if (!query) return true;
    return (m.name && m.name.toLowerCase().includes(query)) ||
           (m.code && m.code.toLowerCase().includes(query)) ||
           (m.formatted_code && m.formatted_code.toLowerCase().includes(query));
  });

  elements.masterTableBody.innerHTML = '';
  list.forEach((m, idx) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="text-center">${idx + 1}</td>
      <td><span class="cell-code">${m.code || '-'}</span></td>
      <td><span class="cell-code">${m.formatted_code || '-'}</span></td>
      <td class="font-bold">${m.name}</td>
      <td class="text-right font-bold text-purple">${m.weight_per_unit.toFixed(2)}</td>
      <td class="text-center">${m.unit || 'กก.'}</td>
      <td class="text-center">
        <button class="btn btn-outline-pea btn-sm" onclick="editMasterItem(${m.id})">แก้ไข</button>
      </td>
    `;
    elements.masterTableBody.appendChild(tr);
  });
}

// Populate Manual Modal Master Select
function populateManualModalSelect() {
  elements.manualMasterSelect.innerHTML = '<option value="">-- เลือกเพื่อกรอกน้ำหนักอัตโนมัติ --</option>';
  state.masterData.forEach(m => {
    const opt = document.createElement('option');
    opt.value = m.id;
    opt.textContent = `${m.name} (${m.weight_per_unit} กก./${m.unit})`;
    elements.manualMasterSelect.appendChild(opt);
  });
}

// Save Manual Item (From Modal)
function handleSaveManualItem() {
  const desc = elements.manualDescInput.value.trim();
  if (!desc) {
    alert('กรุณาระบุชื่อรายการอุปกรณ์');
    return;
  }

  const code = elements.manualCodeInput.value.trim();
  const qty = parseFloat(elements.manualQtyInput.value) || 1;
  const unit = elements.manualUnitInput.value.trim() || 'EA';
  const wt = parseFloat(elements.manualWeightInput.value) || 0;

  const newItem = {
    id: Date.now(),
    code: code,
    code_10: code.replace(/\D/g, ''),
    desc: desc,
    master_name: desc,
    section: 'รายการระบุเพิ่มเติม',
    is_dismantle: true,
    unit: unit,
    qty_estimate: qty,
    qty_issued: qty,
    qty_good_return: 0,
    qty_damaged_return: 0,
    qty_installed: 0,
    calc_qty: qty,
    suggested_qty: qty,
    weight_per_unit: wt,
    total_weight: Math.round((qty * wt) * 1000) / 1000,
    match_type: 'exact',
    selected: true
  };

  state.items.unshift(newItem);
  elements.manualItemModal.style.display = 'none';
  renderCalcTable();
  showToast(`เพิ่มรายการ "${desc}" สำเร็จ`, 'success');
}

// Save Master Item (From Modal)
async function handleSaveMasterItem() {
  const name = elements.masterItemName.value.trim();
  if (!name) {
    alert('กรุณาระบุชื่ออุปกรณ์');
    return;
  }

  const rawCode = elements.masterItemCode.value.trim();
  const codeClean = rawCode.replace(/\D/g, '');
  let formattedCode = '';
  if (codeClean.length === 10) {
    formattedCode = `${codeClean[0]}-${codeClean.slice(1, 3)}-${codeClean.slice(3, 6)}-${codeClean.slice(6)}`;
  }

  const wt = parseFloat(elements.masterItemWeight.value) || 0;
  const unit = elements.masterItemUnit.value.trim() || 'กก.';
  const editId = elements.masterEditId.value;

  if (editId) {
    // Update existing
    const item = state.masterData.find(m => m.id == editId);
    if (item) {
      item.name = name;
      item.code = codeClean;
      item.formatted_code = formattedCode;
      item.weight_per_unit = wt;
      item.unit = unit;
    }
  } else {
    // Add new
    state.masterData.push({
      id: Date.now(),
      code: codeClean,
      formatted_code: formattedCode,
      raw_code: rawCode,
      name: name,
      weight_per_unit: wt,
      unit: unit
    });
  }

  elements.masterItemModal.style.display = 'none';
  renderMasterTable();
  elements.masterCountBadge.textContent = state.masterData.length;
  elements.masterTotalItems.textContent = state.masterData.length;

  // Persist to server
  try {
    await fetch('/api/save-master-data', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(state.masterData)
    });
  } catch (e) {
    console.warn('Could not persist to server:', e);
  }

  showToast('บันทึกข้อมูลอุปกรณ์ในฐานข้อมูลเรียบร้อย', 'success');
}

window.editMasterItem = function(id) {
  const m = state.masterData.find(it => it.id == id);
  if (m) {
    elements.masterModalTitle.textContent = 'แก้ไขข้อมูลอุปกรณ์';
    elements.masterEditId.value = m.id;
    elements.masterItemCode.value = m.code || m.formatted_code || '';
    elements.masterItemName.value = m.name;
    elements.masterItemWeight.value = m.weight_per_unit;
    elements.masterItemUnit.value = m.unit;
    elements.masterItemModal.style.display = 'flex';
  }
};

// Export to Excel
async function handleExportExcel() {
  const selectedItems = state.items.filter(it => it.selected);
  if (selectedItems.length === 0) {
    alert('กรุณาเลือกรายการพัสดุอย่างน้อย 1 รายการก่อนส่งออก Excel');
    return;
  }

  showLoading('กำลังสร้างไฟล์ Excel ตามมาตรฐาน กฟภ....');
  try {
    const payload = {
      metadata: state.metadata,
      items: selectedItems,
      summary: {
        total_kg: selectedItems.reduce((acc, it) => acc + it.total_weight, 0),
        price_per_kg: state.scrapPricePerKg
      }
    };

    const res = await fetch('/api/export-excel', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });

    const data = await res.json();
    if (data.success && (data.data_base64 || data.download_url)) {
      let downloadHref;
      let isBlob = false;

      if (data.data_base64) {
        // Decode base64 to binary Blob directly in browser memory (instant download on Vercel)
        const binaryStr = atob(data.data_base64);
        const len = binaryStr.length;
        const bytes = new Uint8Array(len);
        for (let i = 0; i < len; i++) {
          bytes[i] = binaryStr.charCodeAt(i);
        }
        const blob = new Blob([bytes], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
        downloadHref = URL.createObjectURL(blob);
        isBlob = true;
      } else {
        downloadHref = data.download_url;
      }

      // Trigger automatic browser download
      const a = document.createElement('a');
      a.href = downloadHref;
      a.download = data.filename || 'รายงานคืนเศษเหล็ก_กฟภ.xlsx';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);

      if (isBlob) {
        setTimeout(() => URL.revokeObjectURL(downloadHref), 3000);
      }
      showToast('ดาวน์โหลดไฟล์ Excel เรียบร้อยแล้ว!', 'success');
    } else {
      showToast('เกิดข้อผิดพลาดในการสร้าง Excel: ' + (data.error || 'ไม่สามารถสร้างไฟล์ได้'), 'error');
    }
  } catch (err) {
    showToast('ไม่สามารถส่งคำขอสร้าง Excel: ' + err.message, 'error');
  } finally {
    hideLoading();
  }
}

// Render Printable Report
function renderPrintReport() {
  const selectedItems = state.items.filter(it => it.selected);
  const totalKg = selectedItems.reduce((acc, it) => acc + (it.total_weight || 0), 0);
  const totalTons = totalKg / 1000.0;

  elements.printJobNo.textContent = state.metadata.job_no || '-';
  elements.printJobName.textContent = state.metadata.job_name || '-';
  elements.printPersonName.textContent = state.metadata.person_name || '-';
  elements.printPersonId.textContent = state.metadata.person_id || '-';
  elements.printCostCenter.textContent = state.metadata.cost_center_resp || '-';
  elements.printDateVal.textContent = state.metadata.print_date || new Date().toLocaleDateString('th-TH');
  elements.signPersonName.textContent = `(${state.metadata.person_name || '........................................................'})`;

  elements.printGrandTotalKg.textContent = `${totalKg.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} กก.`;
  elements.printGrandTotalTons.textContent = `${totalTons.toLocaleString('en-US', { minimumFractionDigits: 3, maximumFractionDigits: 3 })} ตัน`;

  elements.printTableBody.innerHTML = '';
  if (selectedItems.length === 0) {
    elements.printTableBody.innerHTML = '<tr><td colspan="8" class="text-center">ไม่มีรายการพัสดุที่เลือกคำนวณ</td></tr>';
  } else {
    selectedItems.forEach((it, idx) => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td class="text-center">${idx + 1}</td>
        <td>${it.code || '-'}</td>
        <td>${it.master_name || it.desc}</td>
        <td class="text-center">${it.unit || 'EA'}</td>
        <td class="text-right">${it.calc_qty.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
        <td class="text-right">${it.weight_per_unit.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
        <td class="text-right font-bold">${it.total_weight.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
        <td>${it.section || ''}</td>
      `;
      elements.printTableBody.appendChild(tr);
    });
  }

  elements.printThaiWordSum.textContent = `รวมรายการพัสดุคืนเศษเหล็กทั้งสิ้น ${selectedItems.length} รายการ | น้ำหนักรวมสุทธิ ${totalKg.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} กิโลกรัม (${totalTons.toFixed(3)} ตัน)`;

  // Apply responsive Fit-to-1-Page & zoom scale
  applyPrintFit();
}

// Apply Print Fit & Zoom Scale
function applyPrintFit() {
  const sheet = elements.printReportSheet || document.getElementById('printReportSheet');
  const toggle = elements.toggleFitOnePage || document.getElementById('toggleFitOnePage');
  const scaleSelect = elements.printScaleSelect || document.getElementById('printScaleSelect');
  if (!sheet) return;

  const isFit = toggle ? toggle.checked : true;
  sheet.classList.toggle('fit-page', isFit);

  const scaleVal = scaleSelect ? scaleSelect.value : 'auto';
  const selectedCount = state.items ? state.items.filter(it => it.selected).length : 0;

  if (scaleVal === 'auto') {
    if (!isFit) {
      sheet.style.transform = 'none';
    } else if (selectedCount <= 12) {
      sheet.style.transform = 'scale(1)';
    } else if (selectedCount <= 16) {
      sheet.style.transform = 'scale(0.93)';
    } else if (selectedCount <= 22) {
      sheet.style.transform = 'scale(0.85)';
    } else if (selectedCount <= 30) {
      sheet.style.transform = 'scale(0.78)';
    } else {
      sheet.style.transform = 'scale(0.72)';
    }
  } else {
    const num = parseFloat(scaleVal);
    if (!isNaN(num) && num > 0) {
      sheet.style.transform = `scale(${num / 100})`;
    } else {
      sheet.style.transform = 'none';
    }
  }
}

// Helpers: Loading Overlay & Toast
function showLoading(text) {
  elements.loadingText.textContent = text || 'กำลังดำเนินการ...';
  elements.loadingOverlay.style.display = 'block';
}

function hideLoading() {
  elements.loadingOverlay.style.display = 'none';
}

function showToast(message, type = 'info') {
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  toast.innerHTML = `
    <span>${type === 'success' ? '✓' : type === 'error' ? '✕' : 'ℹ'}</span>
    <span>${message}</span>
  `;
  elements.toastContainer.appendChild(toast);
  setTimeout(() => {
    toast.style.opacity = '0';
    setTimeout(() => toast.remove(), 300);
  }, 4000);
}

// Start application on DOM ready
document.addEventListener('DOMContentLoaded', init);
