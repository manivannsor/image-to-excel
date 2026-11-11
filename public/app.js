// State Management
const state = {
  apiKey: localStorage.getItem('GEMINI_API_KEY') || '',
  hasServerApiKey: false,
  selectedFiles: [],
  tableData: null,
  currentImageIndex: 0,
  imageUrls: []
};

// DOM Elements
const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
const fileListContainer = document.getElementById('fileListContainer');
const btnProcessAI = document.getElementById('btnProcessAI');
const processStatus = document.getElementById('processStatus');
const btnLoadSample = document.getElementById('btnLoadSample');
const btnExportExcel = document.getElementById('btnExportExcel');
const btnAddRow = document.getElementById('btnAddRow');

const tableBody = document.getElementById('tableBody');
const studentCountBadge = document.getElementById('studentCountBadge');
const genderStatsBadge = document.getElementById('genderStatsBadge');

const tabTableOnly = document.getElementById('tabTableOnly');
const tabSplitView = document.getElementById('tabSplitView');
const imageViewerPane = document.getElementById('imageViewerPane');
const previewImage = document.getElementById('previewImage');
const imageSwitcher = document.getElementById('imageSwitcher');

// Metadata fields
const metaSchoolName = document.getElementById('metaSchoolName');
const metaTitle = document.getElementById('metaTitle');
const metaTeacher = document.getElementById('metaTeacher');
const metaSubject = document.getElementById('metaSubject');
const metaGrade = document.getElementById('metaGrade');
const metaQuarter = document.getElementById('metaQuarter');
const metaDate = document.getElementById('metaDate');

// API Key Modal Elements
const apiKeyModal = document.getElementById('apiKeyModal');
const btnOpenApiKeyModal = document.getElementById('btnOpenApiKeyModal');
const btnCloseApiKeyModal = document.getElementById('btnCloseApiKeyModal');
const btnCancelApiKey = document.getElementById('btnCancelApiKey');
const btnSaveApiKey = document.getElementById('btnSaveApiKey');
const inputApiKey = document.getElementById('inputApiKey');
const apiKeyStatusLabel = document.getElementById('apiKeyStatusLabel');

// Initialize
document.addEventListener('DOMContentLoaded', async () => {
  await checkServerConfig();
  updateApiKeyStatusUI();
  setupEventListeners();
  loadSampleData(); // Load pre-extracted data so user sees working result instantly!
});

async function checkServerConfig() {
  try {
    const res = await fetch('/api/config');
    const data = await res.json();
    state.hasServerApiKey = !!data.hasServerApiKey;
  } catch (e) {
    console.warn('Could not check server config', e);
  }
}

function updateApiKeyStatusUI() {
  if (state.apiKey) {
    apiKeyStatusLabel.innerHTML = `<span class="inline-block w-2 h-2 rounded-full bg-emerald-500 mr-1.5"></span>API Key: ផ្ទាល់ខ្លួន (Custom)`;
  } else if (state.hasServerApiKey) {
    apiKeyStatusLabel.innerHTML = `<span class="inline-block w-2 h-2 rounded-full bg-emerald-500 mr-1.5"></span>API Key: រួចរាល់ (Server)`;
  } else {
    apiKeyStatusLabel.innerHTML = `<span class="inline-block w-2 h-2 rounded-full bg-amber-500 mr-1.5"></span>បញ្ចូល API Key`;
  }
}

function setupEventListeners() {
  // Dropzone drag & drop
  ['dragenter', 'dragover'].forEach(eventName => {
    dropzone.addEventListener(eventName, e => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.add('dragover');
    });
  });

  ['dragleave', 'drop'].forEach(eventName => {
    dropzone.addEventListener(eventName, e => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.remove('dragover');
    });
  });

  dropzone.addEventListener('drop', e => {
    const files = Array.from(e.dataTransfer.files).filter(f => f.type.startsWith('image/'));
    if (files.length > 0) handleSelectedFiles(files);
  });

  dropzone.addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', e => {
    const files = Array.from(e.target.files);
    if (files.length > 0) handleSelectedFiles(files);
  });

  // Action buttons
  btnLoadSample.addEventListener('click', loadSampleData);
  btnProcessAI.addEventListener('click', processWithAI);
  btnExportExcel.addEventListener('click', exportToExcel);
  btnAddRow.addEventListener('click', addNewRow);

  // Tab views
  tabTableOnly.addEventListener('click', () => {
    tabTableOnly.className = 'px-3 py-1 font-semibold rounded-md bg-emerald-50 text-emerald-700 shadow-xs';
    tabSplitView.className = 'px-3 py-1 font-semibold rounded-md text-slate-600 hover:text-slate-900';
    imageViewerPane.classList.add('hidden');
  });

  tabSplitView.addEventListener('click', () => {
    tabSplitView.className = 'px-3 py-1 font-semibold rounded-md bg-emerald-50 text-emerald-700 shadow-xs';
    tabTableOnly.className = 'px-3 py-1 font-semibold rounded-md text-slate-600 hover:text-slate-900';
    imageViewerPane.classList.remove('hidden');
  });

  // API Key Modal
  const serverKeyNotice = document.getElementById('serverKeyNotice');
  const btnClearApiKey = document.getElementById('btnClearApiKey');

  btnOpenApiKeyModal.addEventListener('click', () => {
    inputApiKey.value = state.apiKey;
    if (state.hasServerApiKey) {
      serverKeyNotice.classList.remove('hidden');
    } else {
      serverKeyNotice.classList.add('hidden');
    }
    apiKeyModal.classList.remove('hidden');
  });

  if (btnClearApiKey) {
    btnClearApiKey.addEventListener('click', () => {
      state.apiKey = '';
      localStorage.removeItem('GEMINI_API_KEY');
      inputApiKey.value = '';
      updateApiKeyStatusUI();
      apiKeyModal.classList.add('hidden');
      showToast('បានលុប Key ផ្ទាល់ខ្លួន។ ឥឡូវប្រព័ន្ធប្រើប្រាស់ Server Key!', 'info');
    });
  }

  btnCloseApiKeyModal.addEventListener('click', () => apiKeyModal.classList.add('hidden'));
  btnCancelApiKey.addEventListener('click', () => apiKeyModal.classList.add('hidden'));
  btnSaveApiKey.addEventListener('click', () => {
    const key = inputApiKey.value.trim();
    state.apiKey = key;
    localStorage.setItem('GEMINI_API_KEY', key);
    updateApiKeyStatusUI();
    apiKeyModal.classList.add('hidden');
    showToast('បានរក្សាទុក Gemini API Key ដោយជោគជ័យ!', 'success');
  });
}

function handleSelectedFiles(files) {
  state.selectedFiles = files;
  state.imageUrls = files.map(file => URL.createObjectURL(file));
  renderFileList();
  renderImageSwitcher();
  if (state.imageUrls.length > 0) {
    previewImage.src = state.imageUrls[0];
  }
  showToast(`បានជ្រើសរើស ${files.length} រូបភាព`, 'info');
}

function renderFileList() {
  fileListContainer.innerHTML = '';
  state.selectedFiles.forEach((file, index) => {
    const pill = document.createElement('div');
    pill.className = 'inline-flex items-center px-3 py-1 bg-white border border-slate-200 rounded-lg text-xs shadow-2xs space-x-2';
    pill.innerHTML = `
      <span class="font-semibold text-slate-700">${file.name}</span>
      <span class="text-slate-400">(${(file.size / 1024).toFixed(0)} KB)</span>
      <button class="text-rose-500 hover:text-rose-700 font-bold ml-1" data-index="${index}">&times;</button>
    `;
    pill.querySelector('button').addEventListener('click', (e) => {
      e.stopPropagation();
      state.selectedFiles.splice(index, 1);
      state.imageUrls.splice(index, 1);
      renderFileList();
      renderImageSwitcher();
    });
    fileListContainer.appendChild(pill);
  });
}

function renderImageSwitcher() {
  imageSwitcher.innerHTML = '';
  if (state.imageUrls.length <= 1) return;

  state.imageUrls.forEach((url, idx) => {
    const btn = document.createElement('button');
    btn.className = `px-2 py-0.5 text-[11px] rounded font-semibold border ${idx === state.currentImageIndex ? 'bg-emerald-600 text-white border-emerald-600' : 'bg-white text-slate-600 border-slate-300'}`;
    btn.textContent = `ទំព័រ ${idx + 1}`;
    btn.addEventListener('click', () => {
      state.currentImageIndex = idx;
      previewImage.src = state.imageUrls[idx];
      renderImageSwitcher();
    });
    imageSwitcher.appendChild(btn);
  });
}

async function loadSampleData() {
  try {
    const res = await fetch('/api/sample-data');
    const result = await res.json();
    if (result.success) {
      state.tableData = result.data;
      state.imageUrls = result.sampleImages;
      state.currentImageIndex = 0;
      previewImage.src = state.imageUrls[0];
      renderImageSwitcher();
      renderMetadata(result.data);
      renderTableRows(result.data.rows);
      showToast('បានផ្ទុកទិន្នន័យគំរូ 22 ជួរដោយជោគជ័យ!', 'success');
    }
  } catch (err) {
    console.error('Error loading sample:', err);
    showToast('មិនអាចផ្ទុកទិន្នន័យគំរូបានទេ', 'error');
  }
}

async function processWithAI() {
  if (state.selectedFiles.length === 0) {
    showToast('សូមជ្រើសរើស File រូបភាពជាមុនសិន!', 'warning');
    return;
  }

  if (!state.apiKey && !state.hasServerApiKey) {
    apiKeyModal.classList.remove('hidden');
    showToast('សូមបញ្ចូល Gemini API Key ដើម្បីដំណើរការ', 'warning');
    return;
  }

  const formData = new FormData();
  state.selectedFiles.forEach(file => {
    formData.append('images', file);
  });
  if (state.apiKey) {
    formData.append('apiKey', state.apiKey);
  }

  btnProcessAI.disabled = true;
  btnProcessAI.classList.add('opacity-75', 'cursor-not-allowed');
  processStatus.classList.remove('hidden');
  processStatus.textContent = 'កំពុងដំណើរការវិភាគរូបភាព និងអានអក្សរដោយ AI (Gemini Vision)...';

  try {
    const response = await fetch('/api/extract', {
      method: 'POST',
      body: formData
    });

    const result = await response.json();
    if (!response.ok) {
      throw new Error(result.message || result.error || 'បម្លែងមិនបានជោគជ័យ');
    }

    state.tableData = result.data;
    renderMetadata(result.data);
    renderTableRows(result.data.rows || []);
    showToast('បម្លែងរូបភាពដោយជោគជ័យ!', 'success');
  } catch (err) {
    console.error('Processing error:', err);
    let errorMsg = err.message || 'បម្លែងមិនបានជោគជ័យ';
    try {
      const parsed = JSON.parse(errorMsg);
      if (parsed.error && parsed.error.message) {
        errorMsg = parsed.error.message;
      }
    } catch (e) {}
    showToast(errorMsg, 'error');
  } finally {
    btnProcessAI.disabled = false;
    btnProcessAI.classList.remove('opacity-75', 'cursor-not-allowed');
    processStatus.classList.add('hidden');
  }
}

function renderMetadata(data) {
  if (!data) return;
  metaSchoolName.value = data.schoolNameKh || 'សាលាអន្តរជាតិ វេស្ទើន';
  metaTitle.value = data.title || 'តារាងសម្រង់ពិន្ទុ / Score List';
  metaTeacher.value = data.teacherName || '';
  metaSubject.value = data.subject || '';
  metaGrade.value = data.grade || '';
  metaQuarter.value = data.quarter || '';
  metaDate.value = data.date || '21-Sep-26';
}

function renderTableRows(rows) {
  tableBody.innerHTML = '';
  if (!rows || rows.length === 0) {
    tableBody.innerHTML = `<tr><td colspan="23" class="text-center py-8 text-slate-400">គ្មានទិន្នន័យ (No Data)</td></tr>`;
    updateStats(0, 0, 0);
    return;
  }

  let maleCount = 0;
  let femaleCount = 0;

  rows.forEach((row, index) => {
    const tr = document.createElement('tr');
    tr.className = 'divide-x divide-slate-200';
    tr.dataset.index = index;

    if (row.sex === 'M') maleCount++;
    if (row.sex === 'F') femaleCount++;

    const hw = row.hw || ['', '', '', '', '', ''];
    const qz = row.qz || ['', '', '', '', '', ''];
    const mt = row.mt || ['', '', ''];

    tr.innerHTML = `
      <td class="text-center text-slate-500 font-medium">
        <input class="cell-input text-center" data-field="no" value="${row.no !== undefined ? row.no : index + 1}">
      </td>
      <td>
        <input class="cell-input font-semibold uppercase text-slate-800" data-field="name_en" value="${row.name_en || ''}">
      </td>
      <td>
        <input class="cell-input font-medium text-slate-800" data-field="name_kh" value="${row.name_kh || ''}">
      </td>
      <td>
        <input class="cell-input text-center uppercase" data-field="sex" value="${row.sex || ''}">
      </td>
      <td>
        <input class="cell-input text-center" data-field="group" value="${row.group || ''}">
      </td>
      <td>
        <input class="cell-input text-center" data-field="conduct" value="${row.conduct || ''}">
      </td>
      <td>
        <input class="cell-input text-center" data-field="cp" value="${row.cp || ''}">
      </td>
      <!-- HW 1-6 -->
      <td class="bg-slate-50/60"><input class="cell-input cell-score" data-field="hw-0" value="${hw[0] !== undefined ? hw[0] : ''}"></td>
      <td class="bg-slate-50/60"><input class="cell-input cell-score" data-field="hw-1" value="${hw[1] !== undefined ? hw[1] : ''}"></td>
      <td class="bg-slate-50/60"><input class="cell-input cell-score" data-field="hw-2" value="${hw[2] !== undefined ? hw[2] : ''}"></td>
      <td class="bg-slate-50/60"><input class="cell-input cell-score" data-field="hw-3" value="${hw[3] !== undefined ? hw[3] : ''}"></td>
      <td class="bg-slate-50/60"><input class="cell-input cell-score" data-field="hw-4" value="${hw[4] !== undefined ? hw[4] : ''}"></td>
      <td class="bg-slate-50/60"><input class="cell-input cell-score" data-field="hw-5" value="${hw[5] !== undefined ? hw[5] : ''}"></td>
      <!-- QZ 1-6 -->
      <td><input class="cell-input cell-score" data-field="qz-0" value="${qz[0] !== undefined ? qz[0] : ''}"></td>
      <td><input class="cell-input cell-score" data-field="qz-1" value="${qz[1] !== undefined ? qz[1] : ''}"></td>
      <td><input class="cell-input cell-score" data-field="qz-2" value="${qz[2] !== undefined ? qz[2] : ''}"></td>
      <td><input class="cell-input cell-score" data-field="qz-3" value="${qz[3] !== undefined ? qz[3] : ''}"></td>
      <td><input class="cell-input cell-score" data-field="qz-4" value="${qz[4] !== undefined ? qz[4] : ''}"></td>
      <td><input class="cell-input cell-score" data-field="qz-5" value="${qz[5] !== undefined ? qz[5] : ''}"></td>
      <!-- MT 1-3 -->
      <td class="bg-slate-50/60"><input class="cell-input cell-score" data-field="mt-0" value="${mt[0] !== undefined ? mt[0] : ''}"></td>
      <td class="bg-slate-50/60"><input class="cell-input cell-score" data-field="mt-1" value="${mt[1] !== undefined ? mt[1] : ''}"></td>
      <td class="bg-slate-50/60"><input class="cell-input cell-score" data-field="mt-2" value="${mt[2] !== undefined ? mt[2] : ''}"></td>
      <!-- Delete action -->
      <td class="text-center">
        <button class="text-slate-300 hover:text-rose-500 transition p-1" onclick="deleteRow(${index})">
          <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"></path>
          </svg>
        </button>
      </td>
    `;
    tableBody.appendChild(tr);
  });

  updateStats(rows.length, maleCount, femaleCount);
}

function updateStats(total, male, female) {
  studentCountBadge.textContent = `សិស្សសរុប: ${total} នាក់`;
  genderStatsBadge.textContent = `ស្រី: ${female} | ប្រុស: ${male}`;
}

function addNewRow() {
  if (!state.tableData) {
    state.tableData = { rows: [] };
  }
  const currentCount = state.tableData.rows.length;
  state.tableData.rows.push({
    no: currentCount + 1,
    name_en: '',
    name_kh: '',
    sex: '',
    group: 'III',
    conduct: '',
    cp: '',
    hw: ['', '', '', '', '', ''],
    qz: ['', '', '', '', '', ''],
    mt: ['', '', '']
  });
  renderTableRows(state.tableData.rows);
}

window.deleteRow = function(index) {
  if (!state.tableData || !state.tableData.rows) return;
  state.tableData.rows.splice(index, 1);
  // Re-number
  state.tableData.rows.forEach((r, i) => r.no = i + 1);
  renderTableRows(state.tableData.rows);
};

function collectTableDataFromDOM() {
  const rows = [];
  const trList = tableBody.querySelectorAll('tr');
  trList.forEach(tr => {
    const getVal = field => {
      const input = tr.querySelector(`input[data-field="${field}"]`);
      return input ? input.value.trim() : '';
    };

    const rowObj = {
      no: getVal('no'),
      name_en: getVal('name_en'),
      name_kh: getVal('name_kh'),
      sex: getVal('sex'),
      group: getVal('group'),
      conduct: getVal('conduct'),
      cp: getVal('cp'),
      hw: [getVal('hw-0'), getVal('hw-1'), getVal('hw-2'), getVal('hw-3'), getVal('hw-4'), getVal('hw-5')],
      qz: [getVal('qz-0'), getVal('qz-1'), getVal('qz-2'), getVal('qz-3'), getVal('qz-4'), getVal('qz-5')],
      mt: [getVal('mt-0'), getVal('mt-1'), getVal('mt-2')]
    };

    if (rowObj.name_en || rowObj.name_kh || rowObj.no) {
      rows.push(rowObj);
    }
  });

  return {
    isScoreList: true,
    schoolNameKh: metaSchoolName.value,
    title: metaTitle.value,
    teacherName: metaTeacher.value,
    subject: metaSubject.value,
    grade: metaGrade.value,
    quarter: metaQuarter.value,
    date: metaDate.value,
    rows: rows
  };
}

async function exportToExcel() {
  const currentData = collectTableDataFromDOM();
  if (currentData.rows.length === 0) {
    showToast('គ្មានទិន្នន័យសម្រាប់ទាញយកទេ!', 'warning');
    return;
  }

  btnExportExcel.disabled = true;
  btnExportExcel.classList.add('opacity-75');

  try {
    const filename = `Score_List_${currentData.subject || 'Subject'}_${currentData.grade || '1A'}`;
    const response = await fetch('/api/export', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        data: currentData,
        filename: filename
      })
    });

    if (!response.ok) throw new Error('ការទាញយក Excel បរាជ័យ');

    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${filename}.xlsx`;
    document.body.appendChild(a);
    a.click();
    window.URL.revokeObjectURL(url);
    document.body.removeChild(a);

    showToast('បានទាញយក File Excel (.xlsx) ដោយជោគជ័យ!', 'success');
  } catch (err) {
    console.error('Export error:', err);
    showToast(err.message, 'error');
  } finally {
    btnExportExcel.disabled = false;
    btnExportExcel.classList.remove('opacity-75');
  }
}

// Toast Notifications
function showToast(message, type = 'info') {
  const container = document.getElementById('toastContainer');
  const toast = document.createElement('div');
  const bgColors = {
    success: 'bg-emerald-600 text-white',
    error: 'bg-rose-600 text-white',
    warning: 'bg-amber-500 text-white',
    info: 'bg-slate-800 text-white'
  };

  toast.className = `toast px-4 py-2.5 rounded-xl shadow-lg text-xs font-semibold flex items-center space-x-2 pointer-events-auto ${bgColors[type] || bgColors.info}`;
  toast.innerHTML = `<span>${message}</span>`;
  container.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transition = 'opacity 0.3s ease';
    setTimeout(() => toast.remove(), 300);
  }, 3500);
}

