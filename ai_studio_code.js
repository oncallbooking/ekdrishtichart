// script.js

(function() {
  "use strict";

  // --- App State ---
  const state = {
    rawData: [], // Raw data from the loaded file
    columns: [], // { name: 'colName', type: 'string'|'number'|'date'|'boolean' }
    filteredData: [], // Data after applying filters and global search
    filters: {}, // { 'colName': { type: 'checkbox'|'range'|'dateRange', selected: Set | {min, max} } }
    charts: [], // { id, type, xField, yField, aggregation, chartInstance }
    map: null,
    mapMarkersLayer: null,
    currentMapMarkers: [], // To manage individual markers for fitting bounds
    geocodeCache: {}, // { 'city,state': {lat, lng} }
    pagination: {
      currentPage: 1,
      pageSize: 25,
      sortColumn: null,
      sortDirection: 'asc' // 'asc' or 'desc'
    },
    loadedFileName: null,
    rawWorkbook: null, // Stores XLSX.read result for multi-sheet
    selectedSheet: null, // Current sheet name
  };

  // --- DOM Elements ---
  const uploadDataBtn = document.getElementById('uploadDataBtn');
  const themeToggle = document.getElementById('themeToggle');
  const sidebar = document.getElementById('sidebar');
  const datasetInfo = document.getElementById('datasetInfo');
  const schemaList = document.getElementById('schemaList');
  const filtersContainer = document.getElementById('filtersContainer');
  const resetFiltersBtn = document.getElementById('resetFilters');

  const chartTypeSelect = document.getElementById('chartType');
  const xFieldSelect = document.getElementById('xField');
  const yFieldSelect = document.getElementById('yField');
  const aggSelect = document.getElementById('agg');
  const addChartBtn = document.getElementById('addChart');
  const dashboardGrid = document.getElementById('dashboard');
  const noChartsMessage = document.getElementById('noChartsMessage');

  const tableSearchInput = document.getElementById('tableSearch');
  const dataTableHead = document.getElementById('tableHead');
  const dataTableBody = document.getElementById('tableBody');
  const tableInfo = document.getElementById('tableInfo');
  const prevPageBtn = document.getElementById('prevPage');
  const nextPageBtn = document.getElementById('nextPage');
  const exportTableCSVBtn = document.getElementById('exportTableCSV');

  const renderMapBtn = document.getElementById('renderMapBtn');
  const geocodeAllBtn = document.getElementById('geocodeAllBtn');

  // Upload Modal elements
  const uploadModalEl = document.getElementById('uploadModal');
  const uploadModal = new bootstrap.Modal(uploadModalEl);
  const fileInput = document.getElementById('fileInput');
  const dropZone = document.getElementById('dropZone');
  const chooseFileBtn = document.getElementById('chooseFileBtn');
  const uploadStatus = document.getElementById('uploadStatus');
  const sheetSelectionDiv = document.getElementById('sheetSelection');
  const sheetSelectModal = document.getElementById('sheetSelectModal');
  const previewTableContainerModal = document.getElementById('previewTableContainerModal');
  const previewTableHead = document.getElementById('previewTableHead');
  const previewTableBody = document.getElementById('previewTableBody');
  const loadDataBtn = document.getElementById('loadDataBtn');

  // Full Viz Modal elements
  const fullVizModalEl = document.getElementById('fullVizModal');
  const fullVizModal = new bootstrap.Modal(fullVizModalEl);
  const fullVizModalLabel = document.getElementById('fullVizModalLabel');
  const fullVizContent = document.getElementById('fullVizContent');
  const downloadFullVizBtn = document.getElementById('downloadFullViz');

  // --- Initial Setup ---
  initTheme();
  initMapGlobal(); // Initialize Leaflet map instance once for global use
  Chart.defaults.color = getComputedStyle(document.documentElement).getPropertyValue('--text-color');
  Chart.defaults.borderColor = getComputedStyle(document.documentElement).getPropertyValue('--border-color');
  Chart.defaults.font.family = getComputedStyle(document.body).getPropertyValue('font-family');

  // --- Utility Functions ---

  function generateUniqueId(prefix = 'id') {
    return `${prefix}_${Math.random().toString(36).slice(2, 9)}`;
  }

  function escapeHtml(s) {
    if (s === undefined || s === null) return '';
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#039;');
  }

  function formatDateISO(d) {
    if (!(d instanceof Date) || isNaN(d)) return '';
    return d.toISOString().slice(0, 10);
  }

  function parseDate(dateString) {
    // Attempt multiple common date formats
    const d = new Date(dateString);
    if (!isNaN(d)) return d;

    // YYYY-MM-DD
    if (/^\d{4}-\d{2}-\d{2}$/.test(dateString)) {
      const parts = dateString.split('-');
      return new Date(parts[0], parts[1] - 1, parts[2]);
    }
    // MM/DD/YYYY
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(dateString)) {
      const parts = dateString.split('/');
      return new Date(parts[2], parts[0] - 1, parts[1]);
    }
    // DD-MM-YYYY
    if (/^\d{2}-\d{2}-\d{4}$/.test(dateString)) {
      const parts = dateString.split('-');
      return new Date(parts[2], parts[1] - 1, parts[0]);
    }
    return new Date('Invalid Date');
  }

  // Detect column type from a sample of values
  function detectColumnType(values) {
    const nonNullValues = values.filter(v => v !== null && v !== undefined && v !== '');
    if (nonNullValues.length === 0) return 'string';

    const sample = nonNullValues.slice(0, Math.min(100, nonNullValues.length));

    // Try number
    if (sample.every(v => !isNaN(Number(v)))) return 'number';

    // Try boolean
    if (sample.every(v => typeof v === 'boolean' || ['true', 'false', '0', '1'].includes(String(v).toLowerCase()))) return 'boolean';

    // Try date (more robust check needed)
    if (sample.every(v => !isNaN(parseDate(v)))) return 'date';

    return 'string';
  }

  // --- Theme Management ---
  function initTheme() {
    const savedTheme = localStorage.getItem('dataviz-studio-theme') || 'light';
    document.documentElement.setAttribute('data-theme', savedTheme);
    updateThemeIcon(savedTheme);
  }

  function updateThemeIcon(theme) {
    themeToggle.innerHTML = theme === 'light' ? '<i class="fa-solid fa-moon"></i>' : '<i class="fa-solid fa-sun"></i>';
  }

  // --- Data Loading & Schema ---

  async function parseFile(file) {
    uploadStatus.textContent = `Parsing "${file.name}"...`;
    loadDataBtn.disabled = true;
    state.rawWorkbook = null;
    state.selectedSheet = null;
    previewTableBody.innerHTML = '';
    previewTableHead.innerHTML = '';
    sheetSelectionDiv.style.display = 'none';
    previewTableContainerModal.style.display = 'none';

    try {
      const data = await file.arrayBuffer();
      const fileNameLower = file.name.toLowerCase();
      let workbook;
      let jsonData;

      if (fileNameLower.endsWith('.csv')) {
        const text = new TextDecoder().decode(data);
        jsonData = Papa.parse(text, { header: true, skipEmptyLines: true }).data;
        state.loadedFileName = file.name;
        // Mock a single sheet workbook for CSV consistency
        workbook = { SheetNames: ['Sheet1'], Sheets: { 'Sheet1': XLSX.utils.json_to_sheet(jsonData) } };
      } else { // Assume XLSX/XLS
        workbook = XLSX.read(data, { type: 'array' });
        state.loadedFileName = file.name;
      }
      state.rawWorkbook = workbook;
      populateSheetSelect();
      loadDataBtn.disabled = false;
      uploadStatus.textContent = `File "${file.name}" parsed. Select sheet and load.`;
    } catch (err) {
      console.error("File parsing error:", err);
      uploadStatus.innerHTML = `<span class="text-danger">Error parsing file: ${err.message}</span>`;
      loadDataBtn.disabled = true;
    }
  }

  function populateSheetSelect() {
    if (!state.rawWorkbook || state.rawWorkbook.SheetNames.length === 0) return;

    sheetSelectModal.innerHTML = '';
    state.rawWorkbook.SheetNames.forEach(sheetName => {
      const option = document.createElement('option');
      option.value = sheetName;
      option.textContent = sheetName;
      sheetSelectModal.appendChild(option);
    });

    if (state.rawWorkbook.SheetNames.length > 1) {
      sheetSelectionDiv.style.display = 'block';
    } else {
      sheetSelectionDiv.style.display = 'none';
    }
    // Automatically select the first sheet and trigger preview
    sheetSelectModal.value = state.rawWorkbook.SheetNames[0];
    updatePreviewTable();
  }

  function updatePreviewTable() {
    const sheetName = sheetSelectModal.value;
    if (!state.rawWorkbook || !sheetName) {
      previewTableContainerModal.style.display = 'none';
      return;
    }

    const worksheet = state.rawWorkbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

    if (jsonData.length === 0) {
      previewTableHead.innerHTML = '<tr><td colspan="100" class="text-center text-muted">No data in this sheet.</td></tr>';
      previewTableBody.innerHTML = '';
      previewTableContainerModal.style.display = 'block';
      return;
    }

    const headers = Object.keys(jsonData[0]);
    previewTableHead.innerHTML = `<tr>${headers.map(h => `<th>${escapeHtml(h)}</th>`).join('')}</tr>`;

    previewTableBody.innerHTML = jsonData.slice(0, 10).map(row => {
      return `<tr>${headers.map(h => `<td>${escapeHtml(row[h])}</td>`).join('')}</tr>`;
    }).join('');
    previewTableContainerModal.style.display = 'block';
  }

  function loadDataIntoStudio() {
    const sheetName = sheetSelectModal.value;
    if (!state.rawWorkbook || !sheetName) {
      alert("No sheet selected or workbook not loaded.");
      return;
    }

    const worksheet = state.rawWorkbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

    state.rawData = jsonData;
    state.selectedSheet = sheetName;
    detectSchemaAndBuildUI(jsonData);
    applyFiltersAndRender();
    uploadModal.hide();
    updateDashboardControls();
    datasetInfo.innerHTML = `<strong>File:</strong> ${state.loadedFileName}<br><strong>Sheet:</strong> ${sheetName}<br><strong>Rows:</strong> ${state.rawData.length}`;
  }

  function detectSchemaAndBuildUI(data) {
    if (data.length === 0) {
      state.columns = [];
      schemaList.innerHTML = '<p class="text-muted">No data to infer schema.</p>';
      filtersContainer.innerHTML = '<p class="text-muted small">No data to infer filters.</p>';
      return;
    }

    const sample = data.slice(0, Math.min(200, data.length));
    const cols = Object.keys(sample[0] || {});
    state.columns = cols.map(name => {
      const values = sample.map(r => r[name]);
      return { name, type: detectColumnType(values) };
    });

    buildSchemaUI();
    buildFiltersUI();
    populateChartFieldSelectors();
  }

  function buildSchemaUI() {
    schemaList.innerHTML = '';
    state.columns.forEach(c => {
      const icon = c.type === 'number' ? '<i class="fa-solid fa-hashtag"></i>' :
                   c.type === 'date' ? '<i class="fa-solid fa-calendar-alt"></i>' :
                   c.type === 'boolean' ? '<i class="fa-solid fa-toggle-on"></i>' :
                   '<i class="fa-solid fa-font"></i>';
      schemaList.innerHTML += `<div class="d-flex justify-content-between align-items-center py-1">
        <div><strong>${escapeHtml(c.name)}</strong><div class="small-muted">${c.type}</div></div>
        <div class="small-muted">${icon}</div>
      </div>`;
    });
  }

  // --- Filtering ---

  function buildFiltersUI() {