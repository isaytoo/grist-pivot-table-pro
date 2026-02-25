/**
 * Grist Pivot Table Pro Widget
 * Copyright 2026 Said Hamadou (isaytoo)
 * Licensed under the Apache License, Version 2.0
 * https://github.com/isaytoo/grist-pivot-table-pro
 * 
 * Custom pivot table without data size limitations
 */

// =============================================================================
// TRANSLATIONS
// =============================================================================

var translations = {
  fr: {
    title: 'Pivot Table Pro',
    table: 'Table',
    selectTable: '-- Choisir une table --',
    loading: 'Chargement des données...',
    emptyTitle: 'Aucune donnée',
    emptyDesc: 'Sélectionnez une table source pour commencer.',
    availableFields: 'Champs disponibles',
    rowsZone: 'Lignes',
    colsZone: 'Colonnes',
    valuesZone: 'Valeurs',
    dragFields: 'Glissez des champs',
    dragFieldsDesc: 'Glissez des champs dans les zones ci-dessus pour créer votre tableau croisé dynamique.',
    export: 'Exporter CSV',
    rows: 'Lignes',
    columns: 'Colonnes',
    size: 'Taille',
    total: 'Total',
    grandTotal: 'Total Général',
    sum: 'Somme',
    count: 'Nombre',
    avg: 'Moyenne',
    min: 'Min',
    max: 'Max',
    countDistinct: 'Distinct',
    median: 'Médiane',
    stdev: 'Écart-type',
    variance: 'Variance',
    pctTotal: '% Total',
    pctRow: '% Ligne',
    pctCol: '% Colonne'
  },
  en: {
    title: 'Pivot Table Pro',
    table: 'Table',
    selectTable: '-- Select a table --',
    loading: 'Loading data...',
    emptyTitle: 'No data',
    emptyDesc: 'Select a source table to get started.',
    availableFields: 'Available fields',
    rowsZone: 'Rows',
    colsZone: 'Columns',
    valuesZone: 'Values',
    dragFields: 'Drag fields',
    dragFieldsDesc: 'Drag fields to the zones above to create your pivot table.',
    export: 'Export CSV',
    rows: 'Rows',
    columns: 'Columns',
    size: 'Size',
    total: 'Total',
    grandTotal: 'Grand Total',
    sum: 'Sum',
    count: 'Count',
    avg: 'Average',
    min: 'Min',
    max: 'Max',
    countDistinct: 'Distinct',
    median: 'Median',
    stdev: 'Std Dev',
    variance: 'Variance',
    pctTotal: '% Total',
    pctRow: '% Row',
    pctCol: '% Column'
  }
};

var currentLang = 'fr';

function setLanguage(lang) {
  currentLang = lang;
  document.querySelectorAll('.lang-btn').forEach(function(btn) {
    btn.classList.toggle('active', btn.textContent.trim().toLowerCase() === lang);
  });
  applyTranslations();
}

function t(key) {
  return translations[currentLang][key] || translations['en'][key] || key;
}

function applyTranslations() {
  document.querySelectorAll('[data-i18n]').forEach(function(el) {
    var key = el.getAttribute('data-i18n');
    if (translations[currentLang][key]) {
      el.textContent = translations[currentLang][key];
    }
  });
}

// =============================================================================
// STATE
// =============================================================================

var allTables = [];
var selectedTable = '';
var tableData = [];
var tableColumns = [];
var columnTypes = {};

var pivotConfig = {
  rows: [],
  cols: [],
  values: []
};

// Filters and sorting
var fieldFilters = {};  // { fieldName: [selectedValues] }
var fieldSorts = {};    // { fieldName: 'asc' | 'desc' | null }
var currentFilterField = '';
var currentFilterZone = '';
var allFilterValues = [];
var filteredData = [];

var aggregationFunctions = {
  sum: function(arr) {
    return arr.reduce(function(a, b) { return a + (parseFloat(b) || 0); }, 0);
  },
  count: function(arr) {
    return arr.length;
  },
  avg: function(arr) {
    if (arr.length === 0) return 0;
    var sum = arr.reduce(function(a, b) { return a + (parseFloat(b) || 0); }, 0);
    return sum / arr.length;
  },
  min: function(arr) {
    var nums = arr.map(function(v) { return parseFloat(v) || 0; });
    return Math.min.apply(null, nums);
  },
  max: function(arr) {
    var nums = arr.map(function(v) { return parseFloat(v) || 0; });
    return Math.max.apply(null, nums);
  },
  countDistinct: function(arr) {
    var unique = {};
    arr.forEach(function(v) { unique[String(v)] = true; });
    return Object.keys(unique).length;
  },
  median: function(arr) {
    var nums = arr.map(function(v) { return parseFloat(v) || 0; }).sort(function(a, b) { return a - b; });
    if (nums.length === 0) return 0;
    var mid = Math.floor(nums.length / 2);
    return nums.length % 2 !== 0 ? nums[mid] : (nums[mid - 1] + nums[mid]) / 2;
  },
  stdev: function(arr) {
    if (arr.length === 0) return 0;
    var nums = arr.map(function(v) { return parseFloat(v) || 0; });
    var mean = nums.reduce(function(a, b) { return a + b; }, 0) / nums.length;
    var variance = nums.reduce(function(a, b) { return a + Math.pow(b - mean, 2); }, 0) / nums.length;
    return Math.sqrt(variance);
  },
  variance: function(arr) {
    if (arr.length === 0) return 0;
    var nums = arr.map(function(v) { return parseFloat(v) || 0; });
    var mean = nums.reduce(function(a, b) { return a + b; }, 0) / nums.length;
    return nums.reduce(function(a, b) { return a + Math.pow(b - mean, 2); }, 0) / nums.length;
  }
};

// =============================================================================
// INITIALIZATION
// =============================================================================

grist.ready({
  requiredAccess: 'read table',
  columns: [],  // All columns
  allowSelectBy: false
});

grist.onOptions(function(options) {
  if (options && options.lang) {
    setLanguage(options.lang);
  }
  if (options && options.pivotConfig) {
    pivotConfig = options.pivotConfig;
  }
  if (options && options.displayOptions) {
    displayOptions = options.displayOptions;
  }
  if (options && options.cellFormat) {
    cellFormat = options.cellFormat;
  }
  if (options && options.conditionalRules) {
    conditionalRules = options.conditionalRules;
  }
  if (options && options.fieldFilters) {
    fieldFilters = options.fieldFilters;
  }
  if (options && options.fieldSorts) {
    fieldSorts = options.fieldSorts;
  }
  renderPivotZones();
});

// Use onRecords to get data from the linked table
grist.onRecords(function(records, mappings) {
  if (!records || records.length === 0) {
    showEmptyState();
    return;
  }
  
  // Hide table selector - we use linked data
  document.querySelector('.controls-bar').classList.add('hidden');
  
  // Get columns from first record
  tableColumns = Object.keys(records[0]).filter(function(col) {
    return col !== 'id' && col !== 'manualSort';
  });
  
  // Detect column types
  columnTypes = {};
  tableColumns.forEach(function(col) {
    var sample = records.find(function(r) { 
      return r[col] !== null && r[col] !== undefined && r[col] !== ''; 
    });
    var val = sample ? sample[col] : null;
    if (typeof val === 'number') {
      columnTypes[col] = 'num';
    } else if (typeof val === 'boolean') {
      columnTypes[col] = 'bool';
    } else if (val && !isNaN(Date.parse(val))) {
      columnTypes[col] = 'date';
    } else {
      columnTypes[col] = 'text';
    }
  });
  
  // Store data
  tableData = records;
  selectedTable = 'linked';
  
  // Update stats
  updateStats(records.length, tableColumns.length);
  
  // Render
  renderFieldList();
  renderPivotZones();
  renderPivotTable();
  showMainContent();
});

// Fallback: try to load tables list for full document access mode
tryLoadTables();

async function tryLoadTables() {
  try {
    var tables = await grist.docApi.listTables();
    if (tables && tables.length > 0) {
      allTables = tables;
      
      var select = document.getElementById('table-select');
      select.innerHTML = '<option value="">' + t('selectTable') + '</option>';
      
      tables.forEach(function(tableName) {
        if (!tableName.startsWith('GristHidden_')) {
          var option = document.createElement('option');
          option.value = tableName;
          option.textContent = tableName;
          select.appendChild(option);
        }
      });
      
      // Show table selector
      document.querySelector('.controls-bar').classList.remove('hidden');
    }
  } catch (e) {
    // No full document access - that's OK, we use onRecords
    console.log('Using linked data mode');
  }
}

async function onTableSelect(tableName) {
  selectedTable = tableName;
  
  grist.setOptions({
    pivotConfig: pivotConfig,
    lang: currentLang
  });
  
  if (tableName) {
    await loadTableData(tableName);
  } else {
    showEmptyState();
  }
}

async function loadTableData(tableName) {
  showLoading();
  
  try {
    var columns = await grist.docApi.fetchTable(tableName);
    
    tableColumns = Object.keys(columns).filter(function(col) {
      return col !== 'id' && col !== 'manualSort';
    });
    
    // Detect column types
    columnTypes = {};
    tableColumns.forEach(function(col) {
      var values = columns[col];
      if (values && values.length > 0) {
        var sample = values.find(function(v) { return v !== null && v !== undefined && v !== ''; });
        if (typeof sample === 'number') {
          columnTypes[col] = 'num';
        } else if (typeof sample === 'boolean') {
          columnTypes[col] = 'bool';
        } else if (sample && !isNaN(Date.parse(sample))) {
          columnTypes[col] = 'date';
        } else {
          columnTypes[col] = 'text';
        }
      } else {
        columnTypes[col] = 'text';
      }
    });
    
    // Convert to array of objects
    var rowCount = columns.id ? columns.id.length : 0;
    tableData = [];
    
    for (var i = 0; i < rowCount; i++) {
      var row = { id: columns.id[i] };
      tableColumns.forEach(function(col) {
        row[col] = columns[col] ? columns[col][i] : null;
      });
      tableData.push(row);
    }
    
    // Update stats
    updateStats(rowCount, tableColumns.length);
    
    // Render field list
    renderFieldList();
    
    // Restore pivot config if any
    renderPivotZones();
    
    // Render pivot table
    renderPivotTable();
    
    showMainContent();
    
  } catch (e) {
    console.error('Error loading table data:', e);
    showEmptyState();
  }
}

// =============================================================================
// UI HELPERS
// =============================================================================

function showLoading() {
  document.getElementById('loading-state').classList.remove('hidden');
  document.getElementById('empty-state').classList.add('hidden');
  document.getElementById('main-container').classList.add('hidden');
  document.getElementById('stats-bar').classList.add('hidden');
}

function showEmptyState() {
  document.getElementById('loading-state').classList.add('hidden');
  document.getElementById('empty-state').classList.remove('hidden');
  document.getElementById('main-container').classList.add('hidden');
  document.getElementById('stats-bar').classList.add('hidden');
}

function showMainContent() {
  document.getElementById('loading-state').classList.add('hidden');
  document.getElementById('empty-state').classList.add('hidden');
  document.getElementById('main-container').classList.remove('hidden');
  document.getElementById('stats-bar').classList.remove('hidden');
}

function updateStats(rows, cols) {
  document.getElementById('stat-rows').textContent = rows.toLocaleString();
  document.getElementById('stat-cols').textContent = cols;
  
  // Estimate size
  var sizeBytes = JSON.stringify(tableData).length;
  var sizeKB = (sizeBytes / 1024).toFixed(1);
  var sizeStr = sizeKB > 1024 ? (sizeKB / 1024).toFixed(1) + ' MB' : sizeKB + ' KB';
  document.getElementById('stat-size').textContent = sizeStr;
}

// =============================================================================
// FIELD LIST & DRAG/DROP
// =============================================================================

function renderFieldList() {
  var container = document.getElementById('field-list');
  container.innerHTML = '';
  
  tableColumns.forEach(function(col) {
    var div = document.createElement('div');
    div.className = 'field-item';
    div.draggable = true;
    div.dataset.field = col;
    
    var typeClass = columnTypes[col] || 'text';
    var typeLabel = typeClass === 'num' ? '#' : typeClass === 'date' ? '📅' : typeClass === 'bool' ? '✓' : 'Aa';
    
    div.innerHTML = '<span class="field-type ' + typeClass + '">' + typeLabel + '</span>' +
                    '<span>' + col + '</span>';
    
    div.addEventListener('dragstart', function(e) {
      e.dataTransfer.setData('text/plain', col);
      e.dataTransfer.setData('source', 'field-list');
      div.classList.add('dragging');
    });
    
    div.addEventListener('dragend', function() {
      div.classList.remove('dragging');
    });
    
    container.appendChild(div);
  });
}

function allowDrop(e) {
  e.preventDefault();
  e.currentTarget.classList.add('drag-over');
}

function leaveDrop(e) {
  e.currentTarget.classList.remove('drag-over');
}

function dropField(e, zone) {
  e.preventDefault();
  e.currentTarget.classList.remove('drag-over');
  
  var field = e.dataTransfer.getData('text/plain');
  if (!field) return;
  
  // Remove from other zones first
  ['rows', 'cols', 'values'].forEach(function(z) {
    pivotConfig[z] = pivotConfig[z].filter(function(item) {
      return typeof item === 'string' ? item !== field : item.field !== field;
    });
  });
  
  // Add to target zone
  if (zone === 'values') {
    // Values need aggregation function
    pivotConfig.values.push({
      field: field,
      agg: columnTypes[field] === 'num' ? 'sum' : 'count'
    });
  } else {
    pivotConfig[zone].push(field);
  }
  
  // Save config
  grist.setOptions({
    selectedTable: selectedTable,
    pivotConfig: pivotConfig,
    lang: currentLang
  });
  
  renderPivotZones();
  renderPivotTable();
}

function renderPivotZones() {
  // Rows
  var rowsContainer = document.getElementById('zone-rows-items');
  rowsContainer.innerHTML = '';
  pivotConfig.rows.forEach(function(field) {
    rowsContainer.appendChild(createZoneItem(field, 'rows'));
  });
  
  // Cols
  var colsContainer = document.getElementById('zone-cols-items');
  colsContainer.innerHTML = '';
  pivotConfig.cols.forEach(function(field) {
    colsContainer.appendChild(createZoneItem(field, 'cols'));
  });
  
  // Values
  var valuesContainer = document.getElementById('zone-values-items');
  valuesContainer.innerHTML = '';
  pivotConfig.values.forEach(function(item) {
    valuesContainer.appendChild(createValueZoneItem(item));
  });
}

function createZoneItem(field, zone) {
  var div = document.createElement('div');
  div.className = 'zone-item';
  var hasFilter = fieldFilters[field] && fieldFilters[field].length > 0;
  var filterClass = hasFilter ? 'filter-btn has-filter' : 'filter-btn';
  div.innerHTML = '<span>' + field + '</span>' +
                  '<button class="' + filterClass + '" onclick="openFilterModal(\'' + field + '\', \'' + zone + '\')" title="Filtrer">🔍</button>' +
                  '<button class="remove-btn" onclick="removeFromZone(\'' + field + '\', \'' + zone + '\')">&times;</button>';
  return div;
}

function createValueZoneItem(item) {
  var div = document.createElement('div');
  div.className = 'zone-item';
  
  var aggOptions = ['sum', 'count', 'avg', 'min', 'max', 'countDistinct', 'median', 'stdev', 'variance'];
  var selectHtml = '<select onchange="changeAggregation(\'' + item.field + '\', this.value)">';
  aggOptions.forEach(function(agg) {
    var selected = item.agg === agg ? ' selected' : '';
    selectHtml += '<option value="' + agg + '"' + selected + '>' + t(agg) + '</option>';
  });
  selectHtml += '</select>';
  
  div.innerHTML = '<span>' + item.field + '</span>' +
                  selectHtml +
                  '<button class="remove-btn" onclick="removeFromZone(\'' + item.field + '\', \'values\')">&times;</button>';
  return div;
}

function removeFromZone(field, zone) {
  if (zone === 'values') {
    pivotConfig.values = pivotConfig.values.filter(function(item) {
      return item.field !== field;
    });
  } else {
    pivotConfig[zone] = pivotConfig[zone].filter(function(f) { return f !== field; });
  }
  
  grist.setOptions({
    selectedTable: selectedTable,
    pivotConfig: pivotConfig,
    lang: currentLang
  });
  
  renderPivotZones();
  renderPivotTable();
}

function changeAggregation(field, agg) {
  pivotConfig.values = pivotConfig.values.map(function(item) {
    if (item.field === field) {
      return { field: field, agg: agg };
    }
    return item;
  });
  
  grist.setOptions({
    selectedTable: selectedTable,
    pivotConfig: pivotConfig,
    lang: currentLang
  });
  
  renderPivotTable();
}

// =============================================================================
// PIVOT TABLE RENDERING
// =============================================================================

function renderPivotTable() {
  var placeholder = document.getElementById('pivot-placeholder');
  var table = document.getElementById('pivot-table');
  
  if (pivotConfig.rows.length === 0 && pivotConfig.cols.length === 0 && pivotConfig.values.length === 0) {
    placeholder.classList.remove('hidden');
    table.classList.add('hidden');
    return;
  }
  
  placeholder.classList.add('hidden');
  table.classList.remove('hidden');
  
  // Build pivot data
  var pivotData = buildPivotData();
  
  // Render table
  renderPivotTableHTML(pivotData);
}

function buildPivotData() {
  var rowFields = pivotConfig.rows;
  var colFields = pivotConfig.cols;
  var valueFields = pivotConfig.values;
  
  // Use filtered data
  var data = getFilteredData();
  
  // Get unique values for rows and columns
  var rowKeys = getUniqueKeys(data, rowFields);
  var colKeys = getUniqueKeys(data, colFields);
  
  // Apply sorting to row/col keys
  rowFields.forEach(function(f) {
    if (fieldSorts[f]) {
      rowKeys.sort(function(a, b) {
        if (fieldSorts[f] === 'asc') return a.localeCompare(b);
        return b.localeCompare(a);
      });
    }
  });
  colFields.forEach(function(f) {
    if (fieldSorts[f]) {
      colKeys.sort(function(a, b) {
        if (fieldSorts[f] === 'asc') return a.localeCompare(b);
        return b.localeCompare(a);
      });
    }
  });
  
  // Build aggregation map
  var aggMap = {};
  var rowTotals = {};
  var colTotals = {};
  var grandTotal = {};
  
  // Initialize value accumulators
  valueFields.forEach(function(v) {
    grandTotal[v.field] = [];
  });
  
  data.forEach(function(row) {
    var rowKey = rowFields.map(function(f) { return String(row[f] || ''); }).join('|||');
    var colKey = colFields.map(function(f) { return String(row[f] || ''); }).join('|||');
    var cellKey = rowKey + ':::' + colKey;
    
    if (!aggMap[cellKey]) {
      aggMap[cellKey] = {};
      valueFields.forEach(function(v) {
        aggMap[cellKey][v.field] = [];
      });
    }
    
    if (!rowTotals[rowKey]) {
      rowTotals[rowKey] = {};
      valueFields.forEach(function(v) {
        rowTotals[rowKey][v.field] = [];
      });
    }
    
    if (!colTotals[colKey]) {
      colTotals[colKey] = {};
      valueFields.forEach(function(v) {
        colTotals[colKey][v.field] = [];
      });
    }
    
    valueFields.forEach(function(v) {
      var val = row[v.field];
      aggMap[cellKey][v.field].push(val);
      rowTotals[rowKey][v.field].push(val);
      colTotals[colKey][v.field].push(val);
      grandTotal[v.field].push(val);
    });
  });
  
  return {
    rowKeys: rowKeys,
    colKeys: colKeys,
    aggMap: aggMap,
    rowTotals: rowTotals,
    colTotals: colTotals,
    grandTotal: grandTotal,
    rowFields: rowFields,
    colFields: colFields,
    valueFields: valueFields
  };
}

function getUniqueKeys(data, fields) {
  if (fields.length === 0) return [''];
  
  var keysSet = {};
  data.forEach(function(row) {
    var key = fields.map(function(f) { return String(row[f] || ''); }).join('|||');
    keysSet[key] = true;
  });
  
  return Object.keys(keysSet).sort();
}

function renderPivotTableHTML(pivotData) {
  var thead = document.getElementById('pivot-thead');
  var tbody = document.getElementById('pivot-tbody');
  
  var rowFields = pivotData.rowFields;
  var colFields = pivotData.colFields;
  var valueFields = pivotData.valueFields;
  var rowKeys = pivotData.rowKeys;
  var colKeys = pivotData.colKeys;
  
  // Build header
  var headerHtml = '<tr>';
  
  // Row field headers
  rowFields.forEach(function(f) {
    headerHtml += '<th class="row-header">' + f + '</th>';
  });
  
  // If no row fields, add empty header
  if (rowFields.length === 0) {
    headerHtml += '<th></th>';
  }
  
  // Column headers
  if (colFields.length > 0) {
    colKeys.forEach(function(colKey) {
      var colValues = colKey.split('|||');
      var label = colValues.join(' / ');
      var colspan = valueFields.length || 1;
      headerHtml += '<th class="col-header" colspan="' + colspan + '">' + label + '</th>';
    });
  }
  
  // Value headers if multiple values
  if (valueFields.length > 1 || (valueFields.length === 1 && colFields.length === 0)) {
    valueFields.forEach(function(v) {
      headerHtml += '<th class="col-header">' + v.field + ' (' + t(v.agg) + ')</th>';
    });
  } else if (valueFields.length === 1 && colFields.length > 0) {
    // Already included in col headers
  }
  
  // Total column
  if (colFields.length > 0 && valueFields.length > 0) {
    headerHtml += '<th class="col-header">' + t('total') + '</th>';
  }
  
  headerHtml += '</tr>';
  
  // Sub-header for value fields under each column
  if (colFields.length > 0 && valueFields.length > 1) {
    headerHtml += '<tr>';
    rowFields.forEach(function() {
      headerHtml += '<th></th>';
    });
    if (rowFields.length === 0) {
      headerHtml += '<th></th>';
    }
    colKeys.forEach(function() {
      valueFields.forEach(function(v) {
        headerHtml += '<th class="col-header">' + t(v.agg) + '</th>';
      });
    });
    headerHtml += '<th></th>'; // Total
    headerHtml += '</tr>';
  }
  
  thead.innerHTML = headerHtml;
  
  // Build body
  var bodyHtml = '';
  
  rowKeys.forEach(function(rowKey) {
    bodyHtml += '<tr>';
    
    // Row labels
    var rowValues = rowKey.split('|||');
    if (rowFields.length > 0) {
      rowValues.forEach(function(val) {
        bodyHtml += '<th class="row-header">' + (val || '(vide)') + '</th>';
      });
    } else {
      bodyHtml += '<th>' + t('total') + '</th>';
    }
    
    // Data cells
    var alignStyle = 'text-align:' + cellFormat.align + ';';
    if (colFields.length > 0) {
      colKeys.forEach(function(colKey) {
        var cellKey = rowKey + ':::' + colKey;
        valueFields.forEach(function(v) {
          var values = pivotData.aggMap[cellKey] ? pivotData.aggMap[cellKey][v.field] : [];
          var result = aggregationFunctions[v.agg](values);
          var condStyle = getConditionalStyle(result, v.field);
          bodyHtml += '<td class="value-cell" style="' + alignStyle + condStyle + '">' + formatNumber(result, v.field) + '</td>';
        });
      });
      
      // Row total (only if grandTotals includes rows)
      if (displayOptions.grandTotals !== 'hide' && displayOptions.grandTotals !== 'cols') {
        valueFields.forEach(function(v) {
          var values = pivotData.rowTotals[rowKey] ? pivotData.rowTotals[rowKey][v.field] : [];
          var result = aggregationFunctions[v.agg](values);
          bodyHtml += '<td class="value-cell total-cell" style="' + alignStyle + '">' + formatNumber(result, v.field) + '</td>';
        });
      }
    } else {
      // No columns, just values
      valueFields.forEach(function(v) {
        var values = pivotData.rowTotals[rowKey] ? pivotData.rowTotals[rowKey][v.field] : [];
        var result = aggregationFunctions[v.agg](values);
        var condStyle = getConditionalStyle(result, v.field);
        bodyHtml += '<td class="value-cell" style="' + alignStyle + condStyle + '">' + formatNumber(result, v.field) + '</td>';
      });
    }
    
    bodyHtml += '</tr>';
  });
  
  // Grand total row (only if grandTotals includes cols or show)
  if (rowFields.length > 0 && displayOptions.grandTotals !== 'hide' && displayOptions.grandTotals !== 'rows') {
    var alignStyle = 'text-align:' + cellFormat.align + ';';
    bodyHtml += '<tr class="total-row">';
    bodyHtml += '<th colspan="' + rowFields.length + '">' + t('grandTotal') + '</th>';
    
    if (colFields.length > 0) {
      colKeys.forEach(function(colKey) {
        valueFields.forEach(function(v) {
          var values = pivotData.colTotals[colKey] ? pivotData.colTotals[colKey][v.field] : [];
          var result = aggregationFunctions[v.agg](values);
          bodyHtml += '<td class="value-cell" style="' + alignStyle + '">' + formatNumber(result, v.field) + '</td>';
        });
      });
    }
    
    // Grand total
    if (displayOptions.grandTotals === 'show') {
      valueFields.forEach(function(v) {
        var values = pivotData.grandTotal[v.field] || [];
        var result = aggregationFunctions[v.agg](values);
        bodyHtml += '<td class="value-cell total-cell" style="' + alignStyle + '">' + formatNumber(result, v.field) + '</td>';
      });
    }
    
    bodyHtml += '</tr>';
  }
  
  tbody.innerHTML = bodyHtml;
}

function formatNumber(num, fieldName) {
  if (num === null || num === undefined || isNaN(num)) return '-';
  
  // Apply cell format options
  var decimals = cellFormat.decimals;
  var thousands = cellFormat.thousands;
  var currency = cellFormat.currency;
  var currencyPos = cellFormat.currencyPos;
  
  // Format the number
  var formatted;
  if (Number.isInteger(num) && decimals === 0) {
    formatted = Math.round(num).toString();
  } else {
    formatted = num.toFixed(decimals);
  }
  
  // Add thousands separator
  if (thousands) {
    var parts = formatted.split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, thousands);
    formatted = parts.join(',');
  }
  
  // Add currency symbol
  if (currency) {
    if (currencyPos === 'before') {
      formatted = currency + ' ' + formatted;
    } else {
      formatted = formatted + ' ' + currency;
    }
  }
  
  return formatted;
}

// =============================================================================
// EXPORT
// =============================================================================

function toggleExportMenu() {
  var menu = document.getElementById('export-menu');
  menu.classList.toggle('hidden');
}

// Close export menu when clicking outside
document.addEventListener('click', function(e) {
  var menu = document.getElementById('export-menu');
  var dropdown = document.querySelector('.export-dropdown');
  if (menu && dropdown && !dropdown.contains(e.target)) {
    menu.classList.add('hidden');
  }
});

function exportToCSV() {
  var table = document.getElementById('pivot-table');
  if (table.classList.contains('hidden')) return;
  
  var csv = [];
  var rows = table.querySelectorAll('tr');
  
  rows.forEach(function(row) {
    var rowData = [];
    var cells = row.querySelectorAll('th, td');
    cells.forEach(function(cell) {
      var text = cell.textContent.replace(/"/g, '""');
      rowData.push('"' + text + '"');
    });
    csv.push(rowData.join(','));
  });
  
  var csvContent = csv.join('\n');
  var blob = new Blob(['\ufeff' + csvContent], { type: 'text/csv;charset=utf-8;' });
  var link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = 'pivot_' + selectedTable + '_' + new Date().toISOString().slice(0, 10) + '.csv';
  link.click();
  
  document.getElementById('export-menu').classList.add('hidden');
}

function exportToJSON() {
  var table = document.getElementById('pivot-table');
  if (table.classList.contains('hidden')) return;
  
  var data = [];
  var rows = table.querySelectorAll('tbody tr');
  var headers = [];
  
  // Get headers
  var headerCells = table.querySelectorAll('thead tr:last-child th');
  headerCells.forEach(function(cell) {
    headers.push(cell.textContent.trim());
  });
  
  // Get data rows
  rows.forEach(function(row) {
    var rowData = {};
    var cells = row.querySelectorAll('th, td');
    cells.forEach(function(cell, index) {
      var header = headers[index] || 'col_' + index;
      rowData[header] = cell.textContent.trim();
    });
    data.push(rowData);
  });
  
  var jsonContent = JSON.stringify(data, null, 2);
  var blob = new Blob([jsonContent], { type: 'application/json;charset=utf-8;' });
  var link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = 'pivot_' + selectedTable + '_' + new Date().toISOString().slice(0, 10) + '.json';
  link.click();
  
  document.getElementById('export-menu').classList.add('hidden');
}

function exportToXLSX() {
  var table = document.getElementById('pivot-table');
  if (table.classList.contains('hidden')) return;
  
  // Use SheetJS to export
  var wb = XLSX.utils.book_new();
  var ws = XLSX.utils.table_to_sheet(table);
  XLSX.utils.book_append_sheet(wb, ws, 'Pivot');
  XLSX.writeFile(wb, 'pivot_' + selectedTable + '_' + new Date().toISOString().slice(0, 10) + '.xlsx');
  
  document.getElementById('export-menu').classList.add('hidden');
}

// =============================================================================
// SAVE/LOAD CONFIGURATION
// =============================================================================

var savedConfigs = [];

function loadSavedConfigs() {
  try {
    var stored = localStorage.getItem('pivotTableProConfigs');
    if (stored) {
      savedConfigs = JSON.parse(stored);
    }
  } catch (e) {
    console.error('Error loading configs:', e);
    savedConfigs = [];
  }
}

function saveSavedConfigs() {
  try {
    localStorage.setItem('pivotTableProConfigs', JSON.stringify(savedConfigs));
  } catch (e) {
    console.error('Error saving configs:', e);
  }
}

function renderConfigList() {
  var container = document.getElementById('config-list');
  
  if (savedConfigs.length === 0) {
    container.innerHTML = '<div style="padding:12px;color:#64748b;text-align:center;">Aucune configuration sauvegardée</div>';
    return;
  }
  
  container.innerHTML = '';
  savedConfigs.forEach(function(config, index) {
    var div = document.createElement('div');
    div.className = 'config-item';
    div.innerHTML = '<span>' + config.name + ' <small style="color:#64748b;">(' + config.date + ')</small></span>' +
                    '<div class="config-item-actions">' +
                      '<button class="config-item-btn load" onclick="loadConfig(' + index + ')">Charger</button>' +
                      '<button class="config-item-btn delete" onclick="deleteConfig(' + index + ')">×</button>' +
                    '</div>';
    container.appendChild(div);
  });
}

function saveConfig() {
  var name = document.getElementById('config-name').value.trim();
  if (!name) {
    alert('Veuillez entrer un nom pour la configuration');
    return;
  }
  
  var config = {
    name: name,
    date: new Date().toLocaleDateString(),
    pivotConfig: pivotConfig,
    displayOptions: displayOptions,
    cellFormat: cellFormat,
    conditionalRules: conditionalRules,
    fieldFilters: fieldFilters,
    fieldSorts: fieldSorts
  };
  
  savedConfigs.push(config);
  saveSavedConfigs();
  renderConfigList();
  document.getElementById('config-name').value = '';
}

function loadConfig(index) {
  var config = savedConfigs[index];
  if (!config) return;
  
  pivotConfig = config.pivotConfig || { rows: [], cols: [], values: [] };
  displayOptions = config.displayOptions || displayOptions;
  cellFormat = config.cellFormat || cellFormat;
  conditionalRules = config.conditionalRules || [];
  fieldFilters = config.fieldFilters || {};
  fieldSorts = config.fieldSorts || {};
  
  saveOptions();
  renderPivotZones();
  renderPivotTable();
  closeModal('save-modal');
}

function deleteConfig(index) {
  if (confirm('Supprimer cette configuration ?')) {
    savedConfigs.splice(index, 1);
    saveSavedConfigs();
    renderConfigList();
  }
}

function exportConfig() {
  var config = {
    name: 'Pivot Table Pro Config',
    exportDate: new Date().toISOString(),
    pivotConfig: pivotConfig,
    displayOptions: displayOptions,
    cellFormat: cellFormat,
    conditionalRules: conditionalRules,
    fieldFilters: fieldFilters,
    fieldSorts: fieldSorts
  };
  
  var jsonContent = JSON.stringify(config, null, 2);
  var blob = new Blob([jsonContent], { type: 'application/json;charset=utf-8;' });
  var link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = 'pivot_config_' + new Date().toISOString().slice(0, 10) + '.json';
  link.click();
}

function importConfig(event) {
  var file = event.target.files[0];
  if (!file) return;
  
  var reader = new FileReader();
  reader.onload = function(e) {
    try {
      var config = JSON.parse(e.target.result);
      
      pivotConfig = config.pivotConfig || { rows: [], cols: [], values: [] };
      displayOptions = config.displayOptions || displayOptions;
      cellFormat = config.cellFormat || cellFormat;
      conditionalRules = config.conditionalRules || [];
      fieldFilters = config.fieldFilters || {};
      fieldSorts = config.fieldSorts || {};
      
      saveOptions();
      renderPivotZones();
      renderPivotTable();
      closeModal('save-modal');
      
      alert('Configuration importée avec succès !');
    } catch (err) {
      alert('Erreur lors de l\'import : ' + err.message);
    }
  };
  reader.readAsText(file);
  event.target.value = '';
}

// Load saved configs on init
loadSavedConfigs();

// =============================================================================
// DISPLAY OPTIONS, CELL FORMAT, CONDITIONAL FORMAT
// =============================================================================

var displayOptions = {
  grandTotals: 'show',  // show, hide, rows, cols
  subtotals: 'hide',    // show, hide
  viewMode: 'compact'   // compact, classic, flat
};

var cellFormat = {
  decimals: 2,
  thousands: ' ',
  currency: '',
  currencyPos: 'after',
  align: 'right'
};

var conditionalRules = [];

// Modal functions
function openModal(modalId) {
  document.getElementById(modalId).classList.remove('hidden');
  
  // Populate format field select
  if (modalId === 'format-modal') {
    var select = document.getElementById('format-field');
    select.innerHTML = '<option value="">-- Toutes les valeurs --</option>';
    pivotConfig.values.forEach(function(v) {
      var opt = document.createElement('option');
      opt.value = v.field;
      opt.textContent = v.field;
      select.appendChild(opt);
    });
  }
  
  // Render conditional rules
  if (modalId === 'conditional-modal') {
    renderCondRules();
  }
  
  // Render saved configs
  if (modalId === 'save-modal') {
    renderConfigList();
  }
}

function closeModal(modalId) {
  document.getElementById(modalId).classList.add('hidden');
}

function closeModalOnOverlay(event) {
  if (event.target.classList.contains('modal-overlay')) {
    event.target.classList.add('hidden');
  }
}

function applyDisplayOptions() {
  displayOptions.grandTotals = document.querySelector('input[name="grand-totals"]:checked').value;
  displayOptions.subtotals = document.querySelector('input[name="subtotals"]:checked').value;
  displayOptions.viewMode = document.querySelector('input[name="view-mode"]:checked').value;
  
  saveOptions();
  renderPivotTable();
  closeModal('display-modal');
}

function applyCellFormat() {
  cellFormat.decimals = parseInt(document.getElementById('format-decimals').value) || 2;
  cellFormat.thousands = document.getElementById('format-thousands').value;
  cellFormat.currency = document.getElementById('format-currency').value;
  cellFormat.currencyPos = document.getElementById('format-currency-pos').value;
  cellFormat.align = document.querySelector('input[name="format-align"]:checked').value;
  
  saveOptions();
  renderPivotTable();
  closeModal('format-modal');
}

// Conditional format
function renderCondRules() {
  var container = document.getElementById('cond-rules-container');
  container.innerHTML = '';
  
  conditionalRules.forEach(function(rule, index) {
    var div = document.createElement('div');
    div.className = 'cond-rule';
    div.innerHTML = 
      '<div class="cond-rule-header">' +
        '<span>Règle ' + (index + 1) + '</span>' +
        '<button class="cond-rule-delete" onclick="deleteCondRule(' + index + ')">×</button>' +
      '</div>' +
      '<div class="modal-row">' +
        '<select class="modal-select" onchange="updateCondRule(' + index + ', \'field\', this.value)">' +
          '<option value="">Toutes valeurs</option>' +
          pivotConfig.values.map(function(v) {
            return '<option value="' + v.field + '"' + (rule.field === v.field ? ' selected' : '') + '>' + v.field + '</option>';
          }).join('') +
        '</select>' +
        '<select class="modal-select" onchange="updateCondRule(' + index + ', \'operator\', this.value)">' +
          '<option value="lt"' + (rule.operator === 'lt' ? ' selected' : '') + '>Inférieur à</option>' +
          '<option value="lte"' + (rule.operator === 'lte' ? ' selected' : '') + '>Inférieur ou égal à</option>' +
          '<option value="gt"' + (rule.operator === 'gt' ? ' selected' : '') + '>Supérieur à</option>' +
          '<option value="gte"' + (rule.operator === 'gte' ? ' selected' : '') + '>Supérieur ou égal à</option>' +
          '<option value="eq"' + (rule.operator === 'eq' ? ' selected' : '') + '>Égal à</option>' +
          '<option value="neq"' + (rule.operator === 'neq' ? ' selected' : '') + '>Différent de</option>' +
          '<option value="between"' + (rule.operator === 'between' ? ' selected' : '') + '>Entre</option>' +
        '</select>' +
        '<input type="number" class="modal-input" value="' + (rule.value || 0) + '" onchange="updateCondRule(' + index + ', \'value\', this.value)">' +
      '</div>' +
      '<div class="color-picker-row">' +
        '<label style="font-size:11px;">Couleur texte:</label>' +
        '<input type="color" value="' + (rule.textColor || '#000000') + '" onchange="updateCondRule(' + index + ', \'textColor\', this.value)">' +
        '<label style="font-size:11px;">Fond:</label>' +
        '<input type="color" value="' + (rule.bgColor || '#ffcccc') + '" onchange="updateCondRule(' + index + ', \'bgColor\', this.value)">' +
      '</div>';
    container.appendChild(div);
  });
}

function addCondRule() {
  conditionalRules.push({
    field: '',
    operator: 'lt',
    value: 0,
    textColor: '#000000',
    bgColor: '#ffcccc'
  });
  renderCondRules();
}

function deleteCondRule(index) {
  conditionalRules.splice(index, 1);
  renderCondRules();
}

function updateCondRule(index, prop, value) {
  if (prop === 'value') {
    conditionalRules[index][prop] = parseFloat(value) || 0;
  } else {
    conditionalRules[index][prop] = value;
  }
}

function applyConditionalFormat() {
  saveOptions();
  renderPivotTable();
  closeModal('conditional-modal');
}

function saveOptions() {
  grist.setOptions({
    pivotConfig: pivotConfig,
    lang: currentLang,
    displayOptions: displayOptions,
    cellFormat: cellFormat,
    conditionalRules: conditionalRules,
    fieldFilters: fieldFilters,
    fieldSorts: fieldSorts
  });
}

// Check conditional rules for a value
function getConditionalStyle(value, fieldName) {
  for (var i = 0; i < conditionalRules.length; i++) {
    var rule = conditionalRules[i];
    if (rule.field && rule.field !== fieldName) continue;
    
    var matches = false;
    var num = parseFloat(value) || 0;
    var ruleVal = parseFloat(rule.value) || 0;
    
    switch (rule.operator) {
      case 'lt': matches = num < ruleVal; break;
      case 'lte': matches = num <= ruleVal; break;
      case 'gt': matches = num > ruleVal; break;
      case 'gte': matches = num >= ruleVal; break;
      case 'eq': matches = num === ruleVal; break;
      case 'neq': matches = num !== ruleVal; break;
    }
    
    if (matches) {
      return 'color:' + rule.textColor + ';background:' + rule.bgColor + ';';
    }
  }
  return '';
}

// =============================================================================
// FILTER AND SORT FUNCTIONS
// =============================================================================

function openFilterModal(field, zone) {
  currentFilterField = field;
  currentFilterZone = zone;
  
  // Get unique values for this field
  var valuesSet = {};
  tableData.forEach(function(row) {
    var val = row[field];
    var key = val === null || val === undefined ? '(vide)' : String(val);
    valuesSet[key] = true;
  });
  
  allFilterValues = Object.keys(valuesSet).sort();
  
  // Get currently selected values (or all if no filter)
  var selectedValues = fieldFilters[field] || allFilterValues.slice();
  
  document.getElementById('filter-modal-title').textContent = '🔍 ' + field;
  document.getElementById('filter-search').value = '';
  
  renderFilterValues(allFilterValues, selectedValues);
  updateFilterCount(selectedValues.length, allFilterValues.length);
  
  document.getElementById('filter-modal').classList.remove('hidden');
}

function renderFilterValues(values, selectedValues) {
  var container = document.getElementById('filter-values-list');
  container.innerHTML = '';
  
  values.forEach(function(val) {
    var isChecked = selectedValues.indexOf(val) !== -1;
    var div = document.createElement('label');
    div.className = 'filter-item';
    div.innerHTML = '<input type="checkbox" value="' + val.replace(/"/g, '&quot;') + '"' + (isChecked ? ' checked' : '') + ' onchange="updateFilterSelection()">' +
                    '<span>' + val + '</span>';
    container.appendChild(div);
  });
  
  // Update select all checkbox
  var allChecked = values.length > 0 && values.every(function(v) { return selectedValues.indexOf(v) !== -1; });
  document.getElementById('filter-select-all').checked = allChecked;
}

function updateFilterSelection() {
  var checkboxes = document.querySelectorAll('#filter-values-list input[type="checkbox"]');
  var selected = [];
  checkboxes.forEach(function(cb) {
    if (cb.checked) selected.push(cb.value);
  });
  updateFilterCount(selected.length, allFilterValues.length);
}

function updateFilterCount(selected, total) {
  document.getElementById('filter-count').textContent = selected + ' éléments sur ' + total + ' sélectionnés';
}

function toggleSelectAll(checked) {
  var checkboxes = document.querySelectorAll('#filter-values-list input[type="checkbox"]');
  checkboxes.forEach(function(cb) {
    cb.checked = checked;
  });
  updateFilterCount(checked ? allFilterValues.length : 0, allFilterValues.length);
}

function filterSearchValues(query) {
  var filtered = allFilterValues.filter(function(val) {
    return val.toLowerCase().indexOf(query.toLowerCase()) !== -1;
  });
  var selectedValues = getSelectedFilterValues();
  renderFilterValues(filtered, selectedValues);
}

function getSelectedFilterValues() {
  var checkboxes = document.querySelectorAll('#filter-values-list input[type="checkbox"]');
  var selected = [];
  checkboxes.forEach(function(cb) {
    if (cb.checked) selected.push(cb.value);
  });
  return selected;
}

function sortFilterValues(direction) {
  var sorted = allFilterValues.slice().sort(function(a, b) {
    if (direction === 'asc') return a.localeCompare(b);
    return b.localeCompare(a);
  });
  var selectedValues = getSelectedFilterValues();
  renderFilterValues(sorted, selectedValues);
  
  // Also set field sort
  fieldSorts[currentFilterField] = direction;
}

function selectTop10() {
  // Select only first 10 values
  var checkboxes = document.querySelectorAll('#filter-values-list input[type="checkbox"]');
  var count = 0;
  checkboxes.forEach(function(cb) {
    cb.checked = count < 10;
    count++;
  });
  updateFilterCount(Math.min(10, allFilterValues.length), allFilterValues.length);
}

function applyFilter() {
  var selected = getSelectedFilterValues();
  
  if (selected.length === allFilterValues.length) {
    // All selected = no filter
    delete fieldFilters[currentFilterField];
  } else {
    fieldFilters[currentFilterField] = selected;
  }
  
  saveOptions();
  renderPivotZones();
  renderPivotTable();
  closeModal('filter-modal');
}

function getFilteredData() {
  var data = tableData;
  
  // Apply all field filters
  Object.keys(fieldFilters).forEach(function(field) {
    var allowedValues = fieldFilters[field];
    if (allowedValues && allowedValues.length > 0) {
      data = data.filter(function(row) {
        var val = row[field];
        var key = val === null || val === undefined ? '(vide)' : String(val);
        return allowedValues.indexOf(key) !== -1;
      });
    }
  });
  
  return data;
}

// =============================================================================
// INIT
// =============================================================================

applyTranslations();
