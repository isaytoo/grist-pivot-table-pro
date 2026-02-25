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
  allowSelectBy: false
});

grist.onOptions(function(options) {
  if (options && options.lang) {
    setLanguage(options.lang);
  }
  if (options && options.pivotConfig) {
    pivotConfig = options.pivotConfig;
  }
  if (options && options.selectedTable) {
    selectedTable = options.selectedTable;
    document.getElementById('table-select').value = selectedTable;
    loadTableData(selectedTable);
  }
});

grist.on('message', function(msg) {
  if (msg.tableId) {
    loadTables();
  }
});

loadTables();

async function loadTables() {
  try {
    var tables = await grist.docApi.listTables();
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
    
    if (selectedTable && tables.indexOf(selectedTable) !== -1) {
      select.value = selectedTable;
      loadTableData(selectedTable);
    } else {
      showEmptyState();
    }
  } catch (e) {
    console.error('Error loading tables:', e);
    showEmptyState();
  }
}

async function onTableSelect(tableName) {
  selectedTable = tableName;
  
  grist.setOptions({
    selectedTable: tableName,
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
  div.innerHTML = '<span>' + field + '</span>' +
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
  
  // Get unique values for rows and columns
  var rowKeys = getUniqueKeys(tableData, rowFields);
  var colKeys = getUniqueKeys(tableData, colFields);
  
  // Build aggregation map
  var aggMap = {};
  var rowTotals = {};
  var colTotals = {};
  var grandTotal = {};
  
  // Initialize value accumulators
  valueFields.forEach(function(v) {
    grandTotal[v.field] = [];
  });
  
  tableData.forEach(function(row) {
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
    if (colFields.length > 0) {
      colKeys.forEach(function(colKey) {
        var cellKey = rowKey + ':::' + colKey;
        valueFields.forEach(function(v) {
          var values = pivotData.aggMap[cellKey] ? pivotData.aggMap[cellKey][v.field] : [];
          var result = aggregationFunctions[v.agg](values);
          bodyHtml += '<td class="value-cell">' + formatNumber(result) + '</td>';
        });
      });
      
      // Row total
      valueFields.forEach(function(v) {
        var values = pivotData.rowTotals[rowKey] ? pivotData.rowTotals[rowKey][v.field] : [];
        var result = aggregationFunctions[v.agg](values);
        bodyHtml += '<td class="value-cell total-cell">' + formatNumber(result) + '</td>';
      });
    } else {
      // No columns, just values
      valueFields.forEach(function(v) {
        var values = pivotData.rowTotals[rowKey] ? pivotData.rowTotals[rowKey][v.field] : [];
        var result = aggregationFunctions[v.agg](values);
        bodyHtml += '<td class="value-cell">' + formatNumber(result) + '</td>';
      });
    }
    
    bodyHtml += '</tr>';
  });
  
  // Grand total row
  if (rowFields.length > 0) {
    bodyHtml += '<tr class="total-row">';
    bodyHtml += '<th colspan="' + rowFields.length + '">' + t('grandTotal') + '</th>';
    
    if (colFields.length > 0) {
      colKeys.forEach(function(colKey) {
        valueFields.forEach(function(v) {
          var values = pivotData.colTotals[colKey] ? pivotData.colTotals[colKey][v.field] : [];
          var result = aggregationFunctions[v.agg](values);
          bodyHtml += '<td class="value-cell">' + formatNumber(result) + '</td>';
        });
      });
    }
    
    // Grand total
    valueFields.forEach(function(v) {
      var values = pivotData.grandTotal[v.field] || [];
      var result = aggregationFunctions[v.agg](values);
      bodyHtml += '<td class="value-cell total-cell">' + formatNumber(result) + '</td>';
    });
    
    bodyHtml += '</tr>';
  }
  
  tbody.innerHTML = bodyHtml;
}

function formatNumber(num) {
  if (num === null || num === undefined || isNaN(num)) return '-';
  if (Number.isInteger(num)) return num.toLocaleString();
  return num.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 2 });
}

// =============================================================================
// EXPORT
// =============================================================================

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
}

// =============================================================================
// INIT
// =============================================================================

applyTranslations();
