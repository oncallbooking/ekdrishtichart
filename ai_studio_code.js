// script.js (completed)
// NOTE: This file expects the HTML and CSS you provided (ai_studio_code.html / ai_studio_code.css).
(function() {
  "use strict";

  // --- App State ---
  const state = {
    rawData: [],
    columns: [],
    filteredData: [],
    filters: {}, // { colName: { type: 'checkbox'|'range'|'dateRange', selected: Set | {min,max} } }
    charts: [], // { id, type, xField, yField, aggregation, chartInstance, containerEl }
    map: null,
    mapMarkersLayer: null,
    currentMapMarkers: [],
    geocodeCache: {},
    pagination: {
      currentPage: 1,
      pageSize: 25,
      sortColumn: null,
      sortDirection: 'asc'
    },
    loadedFileName: null,
    rawWorkbook: null,
    selectedSheet: null,
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
  initMapGlobal();
  Chart.defaults.color = getComputedStyle(document.documentElement).getPropertyValue('--text-color') || '#000';
  Chart.defaults.borderColor = getComputedStyle(document.documentElement).getPropertyValue('--border-color') || '#ccc';
  Chart.defaults.font.family = getComputedStyle(document.body).getPropertyValue('font-family') || 'Inter, sans-serif';

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
    const d = new Date(dateString);
    if (!isNaN(d)) return d;
    if (/^\d{4}-\d{2}-\d{2}$/.test(dateString)) {
      const parts = dateString.split('-');
      return new Date(parts[0], parts[1] - 1, parts[2]);
    }
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(dateString)) {
      const parts = dateString.split('/');
      return new Date(parts[2], parts[0] - 1, parts[1]);
    }
    if (/^\d{2}-\d{2}-\d{4}$/.test(dateString)) {
      const parts = dateString.split('-');
      return new Date(parts[2], parts[1] - 1, parts[0]);
    }
    return new Date('Invalid Date');
  }

  function detectColumnType(values) {
    const nonNullValues = values.filter(v => v !== null && v !== undefined && v !== '');
    if (nonNullValues.length === 0) return 'string';
    const sample = nonNullValues.slice(0, Math.min(100, nonNullValues.length));
    if (sample.every(v => !isNaN(Number(v)))) return 'number';
    if (sample.every(v => typeof v === 'boolean' || ['true', 'false', '0', '1'].includes(String(v).toLowerCase()))) return 'boolean';
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

  // --- Map Initialization ---
  function initMapGlobal() {
    try {
      const mapDiv = document.createElement('div');
      mapDiv.id = 'leafletMap';
      mapDiv.style.height = '500px';
      mapDiv.style.width = '100%';
      // don't attach yet; used when rendering map card
      state.map = {
        template: mapDiv,
      };
    } catch (err) {
      console.warn("Leaflet init issue:", err);
      state.map = null;
    }
  }

  // --- Data Loading & Schema (some functions earlier already provided in your file, but we re-hook events here) ---

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
        workbook = { SheetNames: ['Sheet1'], Sheets: { 'Sheet1': XLSX.utils.json_to_sheet(jsonData) } };
      } else {
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

  // --- Filters UI and Logic ---
  function buildFiltersUI() {
    filtersContainer.innerHTML = '';
    state.filters = {}; // reset

    if (!state.columns || state.columns.length === 0) {
      filtersContainer.innerHTML = '<p class="text-muted small">No data to infer filters.</p>';
      return;
    }

    state.columns.forEach(col => {
      const wrapper = document.createElement('div');
      wrapper.className = 'filter-group';
      const title = document.createElement('div');
      title.className = 'd-flex justify-content-between mb-1';
      title.innerHTML = `<div><strong class="small">${escapeHtml(col.name)}</strong></div>`;
      wrapper.appendChild(title);

      // Get distinct values for small cardinality columns
      const colValues = state.rawData.map(r => (r[col.name] === undefined || r[col.name] === null) ? '' : r[col.name]);
      const distinct = Array.from(new Set(colValues)).slice(0, 500); // cap for performance

      if (col.type === 'number') {
        const numeric = colValues.map(v => Number(v)).filter(v => !isNaN(v));
        const min = Math.min(...numeric);
        const max = Math.max(...numeric);
        const idMin = generateUniqueId('rngmin');
        const idMax = generateUniqueId('rngmax');

        wrapper.innerHTML += `
          <div class="small-muted mb-1">Range</div>
          <div class="d-flex gap-2">
            <input id="${idMin}" type="number" class="form-control form-control-sm" placeholder="${min === Infinity ? '' : min}">
            <input id="${idMax}" type="number" class="form-control form-control-sm" placeholder="${max === -Infinity ? '' : max}">
          </div>
        `;
        filtersContainer.appendChild(wrapper);

        state.filters[col.name] = { type: 'range', selected: { min: null, max: null } };

        document.getElementById(idMin).addEventListener('input', (e) => {
          const v = e.target.value;
          state.filters[col.name].selected.min = v === '' ? null : Number(v);
          applyFiltersAndRender();
        });
        document.getElementById(idMax).addEventListener('input', (e) => {
          const v = e.target.value;
          state.filters[col.name].selected.max = v === '' ? null : Number(v);
          applyFiltersAndRender();
        });

      } else if (col.type === 'date') {
        const idFrom = generateUniqueId('datefrom');
        const idTo = generateUniqueId('dateto');
        wrapper.innerHTML += `
          <div class="small-muted mb-1">Date range</div>
          <div class="d-flex gap-2">
            <input id="${idFrom}" type="date" class="form-control form-control-sm">
            <input id="${idTo}" type="date" class="form-control form-control-sm">
          </div>
        `;
        filtersContainer.appendChild(wrapper);

        state.filters[col.name] = { type: 'dateRange', selected: { from: null, to: null } };

        document.getElementById(idFrom).addEventListener('change', (e) => {
          state.filters[col.name].selected.from = e.target.value ? new Date(e.target.value) : null;
          applyFiltersAndRender();
        });
        document.getElementById(idTo).addEventListener('change', (e) => {
          state.filters[col.name].selected.to = e.target.value ? new Date(e.target.value) : null;
          applyFiltersAndRender();
        });

      } else {
        // treat as categorical / string
        // If distinct values are many, show a small search + first 50
        const limited = distinct.slice(0, 50);
        const listId = generateUniqueId('list');
        const searchId = generateUniqueId('search');
        let checkboxesHTML = `<input id="${searchId}" class="form-control form-control-sm mb-1" placeholder="Filter values...">`;
        checkboxesHTML += `<div id="${listId}" class="small-scroll" style="max-height:150px;">`;
        limited.forEach(v => {
          const safeV = escapeHtml(String(v));
          const cbId = generateUniqueId('cb');
          checkboxesHTML += `<div class="form-check"><input class="form-check-input" type="checkbox" value="${safeV}" id="${cbId}"><label class="form-check-label small" for="${cbId}">${safeV}</label></div>`;
        });
        checkboxesHTML += `</div>`;
        wrapper.innerHTML += checkboxesHTML;
        filtersContainer.appendChild(wrapper);

        state.filters[col.name] = { type: 'checkbox', selected: new Set() };

        const listEl = document.getElementById(listId);
        const searchEl = document.getElementById(searchId);

        // checkbox events
        listEl.querySelectorAll('input[type="checkbox"]').forEach(cb => {
          cb.addEventListener('change', (e) => {
            const val = e.target.value;
            if (e.target.checked) state.filters[col.name].selected.add(val);
            else state.filters[col.name].selected.delete(val);
            applyFiltersAndRender();
          });
        });

        // search filter for values
        searchEl.addEventListener('input', (e) => {
          const q = (e.target.value || '').toLowerCase();
          listEl.querySelectorAll('.form-check').forEach(row => {
            const label = row.querySelector('label').textContent || '';
            row.style.display = label.toLowerCase().includes(q) ? '' : 'none';
          });
        });
      }
    });

    // enable reset button
    resetFiltersBtn.disabled = false;
  }

  function resetAllFilters() {
    // reset state.filters selections and inputs
    Object.keys(state.filters).forEach(col => {
      const f = state.filters[col];
      if (f.type === 'range') {
        f.selected = { min: null, max: null };
      } else if (f.type === 'dateRange') {
        f.selected = { from: null, to: null };
      } else if (f.type === 'checkbox') {
        f.selected = new Set();
      }
    });
    // clear DOM inputs quickly by rebuilding UI
    buildFiltersUI();
    applyFiltersAndRender();
  }

  function applyFiltersAndRender() {
    // apply filters and table search
    const globalQ = (tableSearchInput.value || '').toLowerCase().trim();
    const filtered = state.rawData.filter(row => {
      // check each filter
      for (let colName in state.filters) {
        const f = state.filters[colName];
        const val = row[colName];

        if (f.type === 'range') {
          const num = Number(val);
          if (!isNaN(f.selected.min) && f.selected.min !== null) {
            if (isNaN(num) || num < f.selected.min) return false;
          }
          if (!isNaN(f.selected.max) && f.selected.max !== null) {
            if (isNaN(num) || num > f.selected.max) return false;
          }
        } else if (f.type === 'dateRange') {
          const d = val ? parseDate(val) : new Date('Invalid Date');
          if (f.selected.from && (!d || isNaN(d) || d < f.selected.from)) return false;
          if (f.selected.to && (!d || isNaN(d) || d > f.selected.to)) return false;
        } else if (f.type === 'checkbox') {
          if (f.selected.size > 0) {
            const sval = (val === null || val === undefined) ? '' : String(val);
            if (!f.selected.has(sval)) return false;
          }
        }
      }
      // global search across visible columns
      if (globalQ) {
        const values = Object.values(row).map(v => (v === null || v === undefined) ? '' : String(v).toLowerCase());
        if (!values.some(v => v.includes(globalQ))) return false;
      }

      return true;
    });

    state.filteredData = filtered;
    state.pagination.currentPage = 1; // reset to first page when filters change
    renderTable();
    updateDashboardControls();
  }

  // --- Table Rendering & Pagination ---
  function renderTable() {
    // header
    dataTableHead.innerHTML = '';
    dataTableBody.innerHTML = '';

    if (!state.columns || state.columns.length === 0) {
      dataTableBody.innerHTML = '<tr><td colspan="100" class="text-center text-muted py-3">No data loaded.</td></tr>';
      tableInfo.textContent = '';
      prevPageBtn.disabled = true;
      nextPageBtn.disabled = true;
      exportTableCSVBtn.disabled = true;
      return;
    }

    const headers = state.columns.map(c => c.name);
    dataTableHead.innerHTML = headers.map(h => `<th>${escapeHtml(h)}</th>`).join('');

    // sort if requested
    let rows = state.filteredData.slice();
    const { sortColumn, sortDirection, currentPage, pageSize } = state.pagination;
    if (sortColumn) {
      rows.sort((a, b) => {
        const av = a[sortColumn], bv = b[sortColumn];
        if (av === undefined || av === null) return 1;
        if (bv === undefined || bv === null) return -1;
        if (!isNaN(Number(av)) && !isNaN(Number(bv))) {
          return sortDirection === 'asc' ? Number(av) - Number(bv) : Number(bv) - Number(av);
        } else {
          const sa = String(av).toLowerCase(), sb = String(bv).toLowerCase();
          if (sa < sb) return sortDirection === 'asc' ? -1 : 1;
          if (sa > sb) return sortDirection === 'asc' ? 1 : -1;
          return 0;
        }
      });
    }

    const total = rows.length;
    const pages = Math.max(1, Math.ceil(total / pageSize));
    const start = (currentPage - 1) * pageSize;
    const pageRows = rows.slice(start, start + pageSize);

    if (pageRows.length === 0) {
      dataTableBody.innerHTML = '<tr><td colspan="100" class="text-center text-muted py-3">No rows match the current filters.</td></tr>';
    } else {
      dataTableBody.innerHTML = pageRows.map(r => {
        return `<tr>${headers.map(h => `<td>${escapeHtml(r[h])}</td>`).join('')}</tr>`;
      }).join('');
    }

    tableInfo.innerHTML = `<strong>${total}</strong> rows — Page ${currentPage} of ${pages}`;
    prevPageBtn.disabled = currentPage <= 1;
    nextPageBtn.disabled = currentPage >= pages;
    exportTableCSVBtn.disabled = total === 0;
    // allow header sorting by clicking
    dataTableHead.querySelectorAll('th').forEach((th, idx) => {
      th.style.cursor = 'pointer';
      th.onclick = () => {
        const col = headers[idx];
        if (state.pagination.sortColumn === col) {
          state.pagination.sortDirection = state.pagination.sortDirection === 'asc' ? 'desc' : 'asc';
        } else {
          state.pagination.sortColumn = col;
          state.pagination.sortDirection = 'asc';
        }
        renderTable();
      };
    });
  }

  function gotoPrevPage() {
    if (state.pagination.currentPage > 1) {
      state.pagination.currentPage -= 1;
      renderTable();
    }
  }
  function gotoNextPage() {
    const total = state.filteredData.length;
    const pages = Math.max(1, Math.ceil(total / state.pagination.pageSize));
    if (state.pagination.currentPage < pages) {
      state.pagination.currentPage += 1;
      renderTable();
    }
  }

  function exportTableCSV() {
    if (!state.filteredData || state.filteredData.length === 0) {
      alert("No data to export.");
      return;
    }
    const headers = state.columns.map(c => c.name);
    const rows = state.filteredData.map(r => headers.map(h => {
      const v = r[h];
      if (v === null || v === undefined) return '';
      // escape quotes
      return `"${String(v).replace(/"/g, '""')}"`;
    }).join(','));
    const csv = [headers.map(h => `"${h.replace(/"/g, '""')}"`).join(','), ...rows].join('\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    saveAs(blob, (state.loadedFileName ? state.loadedFileName.replace(/\.[^.]+$/, '') : 'data') + '_export.csv');
  }

  // --- Chart Controls ---
  function populateChartFieldSelectors() {
    xFieldSelect.innerHTML = '';
    yFieldSelect.innerHTML = '';
    if (!state.columns) return;
    state.columns.forEach(c => {
      const opt = document.createElement('option');
      opt.value = c.name;
      opt.textContent = `${c.name} (${c.type})`;
      xFieldSelect.appendChild(opt.cloneNode(true));
      yFieldSelect.appendChild(opt.cloneNode(true));
    });
    addChartBtn.disabled = state.columns.length === 0;
  }

  function addChartToDashboard() {
    const chartType = chartTypeSelect.value;
    const xField = xFieldSelect.value;
    const yField = yFieldSelect.value;
    const agg = aggSelect.value;

    if (!xField || !yField) {
      alert("Please choose both X and Y fields.");
      return;
    }

    // prepare aggregated dataset
    const grouped = {};
    state.filteredData.forEach(row => {
      const key = (row[xField] === undefined || row[xField] === null) ? '' : String(row[xField]);
      const val = Number(row[yField]);
      if (!grouped[key]) grouped[key] = [];
      grouped[key].push(val);
    });

    const labels = Object.keys(grouped);
    const data = labels.map(lbl => {
      const arr = grouped[lbl].filter(v => !isNaN(v));
      if (agg === 'count') return grouped[lbl].length;
      if (arr.length === 0) return 0;
      if (agg === 'sum') return arr.reduce((a, b) => a + b, 0);
      if (agg === 'avg') return arr.reduce((a, b) => a + b, 0) / arr.length;
      if (agg === 'min') return Math.min(...arr);
      if (agg === 'max') return Math.max(...arr);
      return 0;
    });

    // create card
    if (noChartsMessage) noChartsMessage.style.display = 'none';
    const card = document.createElement('div');
    card.className = 'card chart-card p-2';
    const header = document.createElement('div');
    header.className = 'd-flex justify-content-between align-items-center mb-2';
    header.innerHTML = `<div><strong>${escapeHtml(chartType.toUpperCase())}</strong><div class="small-muted">${escapeHtml(xField)} vs ${escapeHtml(yField)} &middot; ${escapeHtml(agg)}</div></div>
      <div class="btn-group">
        <button class="btn btn-sm btn-outline-secondary btn-open-full" title="Open full view"><i class="fa-solid fa-expand"></i></button>
        <button class="btn btn-sm btn-outline-danger btn-close-chart" title="Remove chart"><i class="fa-solid fa-trash"></i></button>
      </div>`;
    const canvasContainer = document.createElement('div');
    canvasContainer.className = 'chart-canvas-container';
    const canvas = document.createElement('canvas');
    canvas.width = 800;
    canvas.height = 400;
    canvasContainer.appendChild(canvas);
    card.appendChild(header);
    card.appendChild(canvasContainer);
    dashboardGrid.appendChild(card);

    // Chart configuration mapping
    const config = {
      type: (chartType === 'area' ? 'line' : chartType === 'horizontalBar' ? 'bar' : (chartType === 'doughnut' ? 'doughnut' : (chartType === 'pie' ? 'pie' : chartType))),
      data: {
        labels,
        datasets: [{
          label: `${yField} (${agg})`,
          data,
          tension: chartType === 'line' || chartType === 'area' ? 0.4 : 0,
          fill: chartType === 'area' ? true : false,
        }]
      },
      options: {
        maintainAspectRatio: false,
        responsive: true,
      }
    };

    if (chartType === 'horizontalBar') {
      config.options.indexAxis = 'y';
    }

    const chartInstance = new Chart(canvas.getContext('2d'), config);

    // Save to state
    const chartId = generateUniqueId('chart');
    state.charts.push({ id: chartId, type: chartType, xField, yField, aggregation: agg, chartInstance, containerEl: card });

    // bind buttons
    header.querySelector('.btn-close-chart').addEventListener('click', () => {
      chartInstance.destroy();
      card.remove();
      state.charts = state.charts.filter(c => c.id !== chartId);
      if (state.charts.length === 0 && noChartsMessage) noChartsMessage.style.display = '';
    });

    header.querySelector('.btn-open-full').addEventListener('click', () => {
      fullVizModalLabel.textContent = `${chartType.toUpperCase()} — ${xField} vs ${yField}`;
      // clone canvas to full viz
      fullVizContent.innerHTML = '';
      const fullCanvas = document.createElement('canvas');
      fullCanvas.id = 'chartCanvasFull';
      fullCanvas.width = Math.min(window.innerWidth * 0.9, 1600);
      fullCanvas.height = Math.min(window.innerHeight * 0.8, 900);
      fullVizContent.appendChild(fullCanvas);
      // create a new Chart instance with same data
      const fullConfig = JSON.parse(JSON.stringify(config));
      // ensure colors/loading follow defaults
      new Chart(fullCanvas.getContext('2d'), fullConfig);
      fullVizModal.show();
    });
  }

  // --- Map Rendering (basic) ---
  function renderMap() {
    // look for latitude/longitude columns heuristically
    const latCandidates = state.columns.filter(c => /lat|latitude/i.test(c.name));
    const lngCandidates = state.columns.filter(c => /lon|lng|longitude/i.test(c.name));

    // fallback: maybe there's an 'address' column
    const addrCandidates = state.columns.filter(c => /address|city|location|place/i.test(c.name));

    if (latCandidates.length === 0 || lngCandidates.length === 0) {
      alert("No obvious latitude/longitude columns found. If you have address/city columns, use Auto-Geocode to generate lat/lng.");
      return;
    }

    const latCol = latCandidates[0].name;
    const lngCol = lngCandidates[0].name;

    // create a map card in dashboard
    const existing = document.getElementById('mapCard');
    if (existing) existing.remove();

    const card = document.createElement('div');
    card.className = 'card map-card p-2';
    card.id = 'mapCard';
    const header = document.createElement('div');
    header.className = 'd-flex justify-content-between align-items-center mb-2';
    header.innerHTML = `<div><strong>Map</strong><div class="small-muted">${escapeHtml(latCol)} / ${escapeHtml(lngCol)}</div></div>
      <div><button class="btn btn-sm btn-outline-danger btn-remove-map"><i class="fa-solid fa-trash"></i></button></div>`;
    const mapHolder = document.createElement('div');
    mapHolder.style.height = '420px';
    mapHolder.style.width = '100%';
    mapHolder.id = 'leafletMapCard';
    card.appendChild(header);
    card.appendChild(mapHolder);
    dashboardGrid.appendChild(card);

    header.querySelector('.btn-remove-map').addEventListener('click', () => {
      if (state._leafletInstance) {
        state._leafletInstance.remove();
        state._leafletInstance = null;
      }
      card.remove();
    });

    // initialize leaflet
    const map = L.map('leafletMapCard', { preferCanvas: true }).setView([20, 0], 2);
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      attribution: '&copy; OpenStreetMap contributors'
    }).addTo(map);

    const markers = L.markerClusterGroup();
    state.filteredData.forEach(r => {
      const lat = Number(r[latCol]);
      const lon = Number(r[lngCol]);
      if (!isNaN(lat) && !isNaN(lon)) {
        const m = L.marker([lat, lon]);
        let popup = '<div>';
        state.columns.slice(0, 6).forEach(c => {
          popup += `<div><strong>${escapeHtml(c.name)}:</strong> ${escapeHtml(r[c.name])}</div>`;
        });
        popup += '</div>';
        m.bindPopup(popup);
        markers.addLayer(m);
      }
    });

    map.addLayer(markers);
    if (markers.getLayers().length > 0) {
      map.fitBounds(markers.getBounds(), { maxZoom: 10 });
    }
    state._leafletInstance = map;
  }

  // --- Geocoding via Nominatim (simple, careful about rate limits) ---
  async function geocodeAddressToLatLng(address) {
    if (!address) return null;
    if (state.geocodeCache[address]) return state.geocodeCache[address];

    // Respect Nominatim usage policy: add small delay between requests, identify user agent via headers is not possible here, so be cautious.
    const url = `https://nominatim.openstreetmap.org/search?format=json&q=${encodeURIComponent(address)}&limit=1`;
    try {
      const res = await fetch(url, { headers: { 'Accept-Language': 'en' }});
      if (!res.ok) return null;
      const j = await res.json();
      if (j && j.length > 0) {
        const out = { lat: Number(j[0].lat), lon: Number(j[0].lon) };
        state.geocodeCache[address] = out;
        // small pause to be polite
        await new Promise(r => setTimeout(r, 700));
        return out;
      }
    } catch (err) {
      console.warn("Geocode error", err);
    }
    return null;
  }

  async function geocodeAllAddresses(addressColName, latColName = 'gen_lat', lngColName = 'gen_lng') {
    if (!addressColName) {
      alert("Provide an address/city column name.");
      return;
    }

    // create new columns if not exist
    state.rawData.forEach(r => {
      r[latColName] = r[latColName] || '';
      r[lngColName] = r[lngColName] || '';
    });

    // iterate and geocode missing ones
    for (let i = 0; i < state.rawData.length; i++) {
      const row = state.rawData[i];
      if (row[latColName] && row[lngColName]) continue; // skip if already have
      const addr = row[addressColName];
      if (!addr) continue;
      const g = await geocodeAddressToLatLng(addr);
      if (g) {
        row[latColName] = g.lat;
        row[lngColName] = g.lon;
      }
    }

    // refresh columns (add new lat/lng if absent)
    if (!state.columns.some(c => c.name === latColName)) state.columns.push({ name: latColName, type: 'number' });
    if (!state.columns.some(c => c.name === lngColName)) state.columns.push({ name: lngColName, type: 'number' });

    populateChartFieldSelectors();
    applyFiltersAndRender();
    alert("Geocoding pass complete (may be partial due to rate limits).");
  }

  // --- Dashboard Helpers ---
  function updateDashboardControls() {
    // toggle map/geocode buttons depending on presence of data
    renderMapBtn.disabled = !state.rawData || state.rawData.length === 0;
    geocodeAllBtn.disabled = !state.rawData || state.rawData.length === 0;
  }

  // --- Event Wiring ---
  // upload interactions
  uploadDataBtn.addEventListener('click', () => {
    uploadModal.show();
  });

  chooseFileBtn.addEventListener('click', () => {
    fileInput.click();
  });

  fileInput.addEventListener('change', (e) => {
    const f = e.target.files && e.target.files[0];
    if (f) parseFile(f);
  });

  // drag/drop
  ['dragenter', 'dragover'].forEach(ev => {
    dropZone.addEventListener(ev, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.add('dragover');
    });
  });
  ['dragleave', 'drop'].forEach(ev => {
    dropZone.addEventListener(ev, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.remove('dragover');
    });
  });
  dropZone.addEventListener('drop', (e) => {
    const f = (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]);
    if (f) parseFile(f);
  });

  sheetSelectModal.addEventListener('change', updatePreviewTable);
  loadDataBtn.addEventListener('click', loadDataIntoStudio);

  resetFiltersBtn.addEventListener('click', () => {
    resetAllFilters();
  });

  tableSearchInput.addEventListener('input', debounce(() => {
    applyFiltersAndRender();
  }, 300));

  prevPageBtn.addEventListener('click', gotoPrevPage);
  nextPageBtn.addEventListener('click', gotoNextPage);
  exportTableCSVBtn.addEventListener('click', exportTableCSV);

  addChartBtn.addEventListener('click', addChartToDashboard);

  chartTypeSelect.addEventListener('change', () => {
    // for scatter/bubble we might want numeric X/Y, but leave to user
  });

  renderMapBtn.addEventListener('click', renderMap);
  geocodeAllBtn.addEventListener('click', async () => {
    // ask user which column is address-like
    const addrCols = state.columns.filter(c => /address|city|location|place/i.test(c.name)).map(c => c.name);
    let chosen = null;
    if (addrCols.length === 1) chosen = addrCols[0];
    else if (addrCols.length > 1) {
      chosen = prompt("Which column contains address/city/place to geocode? Candidates:\n" + addrCols.join(', '), addrCols[0]);
    } else {
      chosen = prompt("No obvious address/city column detected. Enter the column name that contains address/city/place:");
    }
    if (chosen) {
      await geocodeAllAddresses(chosen);
    }
  });

  // full viz download
  downloadFullVizBtn.addEventListener('click', () => {
    // attempt to find a canvas inside fullVizContent and download as PNG
    const c = fullVizContent.querySelector('canvas');
    if (!c) {
      alert("No canvas found to download.");
      return;
    }
    const url = c.toDataURL('image/png');
    const a = document.createElement('a');
    a.href = url;
    a.download = 'visualization.png';
    a.click();
  });

  // theme toggle
  themeToggle.addEventListener('click', () => {
    const cur = document.documentElement.getAttribute('data-theme') || 'light';
    const next = cur === 'light' ? 'dark' : 'light';
    document.documentElement.setAttribute('data-theme', next);
    localStorage.setItem('dataviz-studio-theme', next);
    updateThemeIcon(next);
  });

  // simple debounce helper
  function debounce(fn, ms = 250) {
    let t;
    return (...args) => {
      clearTimeout(t);
      t = setTimeout(() => fn(...args), ms);
    };
  }

  // --- Final initialization tweaks ---
  // start with no charts message visible
  if (state.charts.length === 0 && noChartsMessage) noChartsMessage.style.display = '';

})();
