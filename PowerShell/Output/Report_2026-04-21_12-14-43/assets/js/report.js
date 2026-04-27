/**
 * SharePoint Permission Analyzer - Główny moduł JavaScript raportu
 * Wersja: 1.0.0
 *
 * Odpowiada za:
 * - Ładowanie i przetwarzanie danych z data.js
 * - Dashboard (statystyki, wykresy)
 * - Tabela uprawnień (DataTables, filtry, sortowanie)
 * - Drzewo lokalizacji (jsTree)
 * - Wyszukiwanie użytkownika/grupy
 * - Remediacja (generowanie PS1, eksport planu)
 * - Zabezpieczenia XSS (escapeHtml)
 * - Eksport CSV/JSON
 *
 * BEZPIECZEŃSTWO: Wszystkie dane wejściowe renderowane przez escapeHtml().
 * Brak żadnych poświadczeń ani sekretów w tym pliku.
 */

'use strict';

/* ============================================================
   SEKCJA 1: NARZĘDZIA I BEZPIECZEŃSTWO
   ============================================================ */

/**
 * Escaping HTML - ochrona przed XSS.
 * Używaj ZAWSZE przy wstawianiu danych z data.js do innerHTML.
 * @param {*} str
 * @returns {string}
 */
function escapeHtml(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/**
 * Bezpieczne ustawianie textContent (nie innerHTML).
 * @param {HTMLElement} el
 * @param {*} text
 */
function setTextSafe(el, text) {
  if (el) el.textContent = (text === null || text === undefined) ? '' : String(text);
}

/**
 * Formatowanie liczby ze spacją jako separatorem tysięcy.
 * @param {number} n
 * @returns {string}
 */
function formatNumber(n) {
  if (n === null || n === undefined) return '0';
  return Number(n).toLocaleString('pl-PL');
}

/**
 * Skracanie długich URL-i do wyświetlenia.
 * @param {string} url
 * @param {number} maxLen
 * @returns {string}
 */
function truncateUrl(url, maxLen) {
  maxLen = maxLen || 60;
  if (!url) return '';
  if (url.length <= maxLen) return url;
  var half = Math.floor((maxLen - 3) / 2);
  return url.substring(0, half) + '...' + url.substring(url.length - half);
}

/* ============================================================
   SEKCJA 2: INICJALIZACJA
   ============================================================ */

var App = {
  data: null,           // Pełne dane z data.js (window.SCAN_DATA)
  flatRows: [],         // Spłaszczone wiersze (obiekt + przypisanie)
  filteredRows: [],     // Wiersze po zastosowaniu filtrów
  selectedRows: [],     // Zaznaczone wiersze do remediacji
  dataTable: null,      // Instancja DataTables
  charts: {},           // Instancja wykresów Chart.js
  treeInstance: null,   // Instancja jsTree
  currentTheme: 'light',
  detailModalObject: null,  // Aktualny obiekt w modalu szczegółów
  allPrincipals: [],    // Deduplikowana lista principalów do autocomplete
  allWebApps: [],
  allSiteCollections: [],
  allPermLevels: [],
  objectMap: {},
  urlIndex: {},
  resolvedParentMap: {},
  childrenMap: {},
};

// Stan kontekstu SharePoint REST API
var SpContext = {
  detected: false,
  siteUrl: '',
  connected: false,
  principalCache: {},
  digestCache: {},
  connectionPromise: null,
  autoConnectAttempted: false
};

document.addEventListener('DOMContentLoaded', function() {
  // Sprawdź czy dane zostały załadowane
  if (typeof window.SCAN_DATA === 'undefined' || window.SCAN_DATA === null) {
    showLoadError(window.SCAN_DATA_ERROR || 'Plik data.js nie został załadowany lub jest pusty.');
    return;
  }

  App.data = window.SCAN_DATA;
  buildObjectIndexes();
  initTheme();
  buildFlatRows();
  buildGroupedRows();
  buildUniqueListsForFilters();
  showUI();
  initDashboard();
  initTable();
  initFilters();
  initTree();
  initUserLookup();
  initRemediationTab();
  initChangelogTab();
  initNavigation();
  initThemeToggle();
  updateReportHeader();
});

function showLoadError(msg) {
  var errorEl = document.getElementById('loadError');
  var detailEl = document.getElementById('errorDetail');
  if (errorEl) errorEl.style.display = 'flex';
  if (detailEl && msg) detailEl.textContent = msg;
}

function showUI() {
  var nav = document.getElementById('mainNav');
  var main = document.getElementById('appMain');
  if (nav) nav.style.display = '';
  if (main) main.style.display = '';
}

/* ============================================================
   SEKCJA 3: PRZETWARZANIE DANYCH
   ============================================================ */

/**
 * Spłaszcza drzewo obiektów do tablicy wierszy.
 * Każdy wiersz = jeden obiekt + jedno przypisanie (lub obiekt bez przypisań).
 */
function buildFlatRows() {
  App.flatRows = [];

  if (!App.data || !App.data.Objects) return;

  App.data.Objects.forEach(function(obj) {
    if (!obj.Assignments || obj.Assignments.length === 0) {
      // Obiekt bez przypisań (tylko dziedziczenie)
      App.flatRows.push({
        obj: obj,
        assignment: null,
        isInherited: !obj.HasUniquePermissions,
        rowId: obj.ObjectId + '_empty'
      });
    } else {
      obj.Assignments.forEach(function(asgn, idx) {
        App.flatRows.push({
          obj: obj,
          assignment: asgn,
          isInherited: !obj.HasUniquePermissions,
          rowId: obj.ObjectId + '_' + idx
        });
      });
    }
  });

  App.filteredRows = App.flatRows.slice();
}

function buildGroupedRows() {
  var groupOrder = [];
  var groupMap = {};
  App.flatRows.forEach(function(row) {
    var gid = row.obj.ObjectId;
    if (!groupMap[gid]) {
      groupMap[gid] = { groupId: gid, obj: row.obj, assignments: [] };
      groupOrder.push(gid);
    }
    if (row.assignment) {
      groupMap[gid].assignments.push({ assignment: row.assignment, rowId: row.rowId });
    }
  });
  App.groupedRows = groupOrder.map(function(gid) { return groupMap[gid]; });
  App.groupedRowMap = groupMap;
  App.filteredGroupedRows = App.groupedRows.slice();
}

function buildObjectIndexes() {
  App.objectMap = {};
  App.urlIndex = {};
  App.resolvedParentMap = {};
  App.childrenMap = {};

  if (!App.data || !App.data.Objects) return;

  App.data.Objects.forEach(function(obj) {
    App.objectMap[obj.ObjectId] = obj;
    var url = normalizeTreeUrl(obj.ServerRelativeUrl);
    var webApp = (obj.WebApplicationUrl || '').toLowerCase().replace(/\/+$/, '');
    if (url) App.urlIndex[webApp + '|' + url] = obj.ObjectId;
  });

  App.data.Objects.forEach(function(obj) {
    var parentObjectId = resolveParentObjectId(obj, App.urlIndex);
    App.resolvedParentMap[obj.ObjectId] = parentObjectId || '';

    if (!parentObjectId) return;
    if (!App.childrenMap[parentObjectId]) App.childrenMap[parentObjectId] = [];
    App.childrenMap[parentObjectId].push(obj);
  });
}

function resolveParentObjectId(obj, urlIndex) {
  var parentObjectId = obj.ParentObjectId || '';
  var itemTypes = { 'File': true, 'Folder': true, 'ListItem': true, 'Web': true };

  if (itemTypes[obj.ObjectType] && obj.ServerRelativeUrl) {
    var itemUrl = normalizeTreeUrl(obj.ServerRelativeUrl);
    if (itemUrl) {
      var lastSlash = itemUrl.lastIndexOf('/');
      if (lastSlash > 0) {
        var parentUrl = itemUrl.substring(0, lastSlash);
        var webApp = (obj.WebApplicationUrl || '').toLowerCase().replace(/\/+$/, '');
        var resolved = (urlIndex || App.urlIndex || {})[webApp + '|' + parentUrl];
        if (resolved && resolved !== obj.ObjectId) {
          parentObjectId = resolved;
        }
      }
    }
  }

  return parentObjectId;
}

function getObjectUrl(obj) {
  if (!obj) return '';

  var fullUrl = obj.FullUrl || '';
  if (fullUrl) {
    if (/^[a-z]+:\/\//i.test(fullUrl)) return fullUrl;
    var fullBase = obj.WebUrl || obj.SiteCollectionUrl || obj.WebApplicationUrl || window.location.href;
    try {
      return new URL(fullUrl, fullBase.replace(/\/?$/, '/')).href;
    } catch (err) {}
  }

  if (obj.ServerRelativeUrl) {
    var rootBase = obj.WebApplicationUrl || obj.SiteCollectionUrl || obj.WebUrl || window.location.origin || '';
    try {
      return new URL(obj.ServerRelativeUrl, rootBase).href;
    } catch (err2) {}
  }

  return fullUrl || obj.ServerRelativeUrl || obj.SiteCollectionUrl || obj.WebApplicationUrl || '';
}

function getDirectChildren(objectId) {
  var children = (App.childrenMap && App.childrenMap[objectId]) || [];
  return children.slice().sort(compareObjectsByDisplay);
}

function getDescendantObjects(objectId) {
  var descendants = [];
  var queue = getDirectChildren(objectId);

  while (queue.length > 0) {
    var child = queue.shift();
    descendants.push(child);
    var next = getDirectChildren(child.ObjectId);
    if (next.length > 0) {
      Array.prototype.push.apply(queue, next);
    }
  }

  return descendants;
}

function compareObjectsByDisplay(a, b) {
  var order = {
    'SiteCollection': 1,
    'Web': 2,
    'List': 3,
    'Library': 4,
    'Folder': 5,
    'File': 6,
    'ListItem': 7
  };
  var aRank = order[a.ObjectType] || 99;
  var bRank = order[b.ObjectType] || 99;
  if (aRank !== bRank) return aRank - bRank;

  var aName = (a.Title || a.Name || a.FileLeafRef || a.ObjectId || '').toLowerCase();
  var bName = (b.Title || b.Name || b.FileLeafRef || b.ObjectId || '').toLowerCase();
  return aName.localeCompare(bName);
}

function shouldShowContentSection(obj) {
  return ['SiteCollection', 'Web', 'List', 'Library', 'Folder'].indexOf(obj.ObjectType) !== -1;
}

function buildObjectContentSection(obj, compact) {
  if (!obj || !shouldShowContentSection(obj)) return '';

  var directChildren = getDirectChildren(obj.ObjectId);
  var descendants = getDescendantObjects(obj.ObjectId);
  var stats = { Web: 0, List: 0, Library: 0, Folder: 0, File: 0, ListItem: 0 };

  descendants.forEach(function(child) {
    if (stats[child.ObjectType] !== undefined) {
      stats[child.ObjectType]++;
    }
  });

  var html = '<div class="detail-section-title">Zawartość obiektu</div>';

  if (descendants.length === 0) {
    html += '<p class="info-text text-muted">Brak obiektów potomnych w danych skanu.</p>';
    return html;
  }

  var summaryParts = [
    'Bezpośrednio: <strong>' + directChildren.length + '</strong>',
    'Foldery: <strong>' + stats.Folder + '</strong>',
    'Pliki: <strong>' + stats.File + '</strong>'
  ];

  if (stats.ListItem > 0) summaryParts.push('Elementy list: <strong>' + stats.ListItem + '</strong>');
  if (stats.Library > 0) summaryParts.push('Biblioteki: <strong>' + stats.Library + '</strong>');
  if (stats.List > 0) summaryParts.push('Listy: <strong>' + stats.List + '</strong>');
  if (stats.Web > 0) summaryParts.push('Witryny: <strong>' + stats.Web + '</strong>');

  html += '<p class="info-text">' + summaryParts.join(' | ') + '</p>';

  if (directChildren.length === 0) {
    html += '<p class="info-text text-muted">Brak bezpośrednich dzieci, ale obiekt ma głębszych potomków w danych skanu.</p>';
    return html;
  }

  var limit = compact ? 8 : 15;
  html += '<table class="detail-table">';
  html += '<thead><tr><th>Typ</th><th>Nazwa</th><th>URL</th><th>ACL</th></tr></thead><tbody>';

  directChildren.slice(0, limit).forEach(function(child) {
    var childUrl = getObjectUrl(child);
    var childName = child.Title || child.Name || child.FileLeafRef || child.ObjectId;
    html += '<tr>';
    html += '<td>' + getObjectTypeIcon(child.ObjectType) + ' ' + escapeHtml(child.ObjectType || '') + '</td>';
    html += '<td><a href="#" onclick="showObjectDetail(\'' + escapeHtml(child.ObjectId) + '\'); return false;">' + escapeHtml(childName) + '</a></td>';
    html += '<td>' + (childUrl ? '<a href="' + escapeHtml(childUrl) + '" target="_blank" rel="noopener">' + escapeHtml(truncateUrl(childUrl, compact ? 36 : 60)) + '</a>' : '<span class="text-muted">-</span>') + '</td>';
    html += '<td>' + (child.HasUniquePermissions ? '<span class="badge badge-unique">Unikatowe</span>' : '<span class="badge badge-inherited">Dziedziczone</span>') + '</td>';
    html += '</tr>';
  });

  if (directChildren.length > limit) {
    html += '<tr><td colspan="4" class="text-muted">... i jeszcze ' + (directChildren.length - limit) + ' obiektów bezpośrednich</td></tr>';
  }

  html += '</tbody></table>';
  return html;
}
function buildUniqueListsForFilters() {
  var webApps = {};
  var siteColls = {};
  var permLevels = {};
  var principals = {};

  App.flatRows.forEach(function(row) {
    var obj = row.obj;
    if (obj.WebApplicationUrl) webApps[obj.WebApplicationUrl] = true;
    if (obj.SiteCollectionUrl) siteColls[obj.SiteCollectionUrl] = true;

    var asgn = row.assignment;
    if (asgn) {
      if (asgn.PermissionLevels) {
        asgn.PermissionLevels.forEach(function(pl) {
          if (pl) permLevels[pl] = true;
        });
      }
      var key = asgn.LoginName || asgn.DisplayName;
      if (key && !principals[key]) {
        principals[key] = {
          loginName: asgn.LoginName || '',
          displayName: asgn.DisplayName || '',
          email: asgn.Email || '',
          principalType: asgn.PrincipalType || ''
        };
      }
    }
  });

  App.allWebApps = Object.keys(webApps).sort();
  App.allSiteCollections = Object.keys(siteColls).sort();
  App.allPermLevels = Object.keys(permLevels).sort();
  App.allPrincipals = Object.values(principals);
}

/* ============================================================
   SEKCJA 4: DASHBOARD
   ============================================================ */

function initDashboard() {
  renderStatsCards();
  renderCharts();
  renderTopUniqueObjects();
  renderAlerts();
}

function renderStatsCards() {
  var stats = App.data.Statistics || {};
  var meta = App.data.ScanMetadata || {};

  var durationSec = meta.ScanDuration || 0;
  var durationStr = formatDuration(durationSec);

  function makeCardHtml(c) {
    var val = typeof c.value === 'number' ? formatNumber(c.value) : c.value;
    return '<div class="stat-card">' +
      '<div class="stat-card-icon ' + escapeHtml(c.cls) + '">' +
      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' + c.svgPath + '</svg>' +
      '</div>' +
      '<div class="stat-card-value">' + escapeHtml(val) + '</div>' +
      '<div class="stat-card-label">' + escapeHtml(c.label) + '</div>' +
      '</div>';
  }

  // Grupa 1: Hierarchia farmy (4 poziomy)
  var hierarchyCards = [
    { cls: 'icon-webapp', value: stats.WebApplicationCount || 0, label: 'Aplikacje Web',
      svgPath: '<rect x="2" y="3" width="20" height="14" rx="2"/><line x1="8" y1="21" x2="16" y2="21"/><line x1="12" y1="17" x2="12" y2="21"/>' },
    { cls: 'icon-site', value: stats.SiteCollectionCount || 0, label: 'Site Collections',
      svgPath: '<path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>' },
    { cls: 'icon-web', value: stats.WebCount || 0, label: 'Witryny (Webs)',
      svgPath: '<circle cx="12" cy="12" r="10"/><line x1="2" y1="12" x2="22" y2="12"/><path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"/>' },
    { cls: 'icon-list', value: stats.ListCount || 0, label: 'Listy i biblioteki',
      svgPath: '<line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/>' }
  ];

  // Grupa 2: Metryki skanowania
  var metricCards = [
    { cls: 'icon-file', value: (stats.ItemCount || 0) + (stats.FolderCount || 0), label: 'Pliki, foldery, elementy',
      svgPath: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/>' },
    { cls: 'icon-unique', value: stats.UniquePermissionsCount || 0, label: 'Obiekty z unikat. ACL',
      svgPath: '<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>' },
    { cls: 'icon-assign', value: stats.TotalAssignments || 0, label: 'Przypisania uprawnień',
      svgPath: '<path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/>' },
    { cls: 'icon-error', value: stats.ErrorCount || 0, label: 'Błędy skanowania',
      svgPath: '<circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>' },
    { cls: 'icon-web', value: stats.TotalObjectsScanned || 0, label: 'Łącznie przeskanowanych',
      svgPath: '<polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/>' },
    { cls: 'icon-list', value: durationStr, label: 'Czas skanowania',
      svgPath: '<circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/>' }
  ];

  var gridH = document.getElementById('statsGridHierarchy');
  if (gridH) gridH.innerHTML = hierarchyCards.map(makeCardHtml).join('');

  var gridM = document.getElementById('statsGridMetrics');
  if (gridM) gridM.innerHTML = metricCards.map(makeCardHtml).join('');
}

function formatDuration(seconds) {
  if (!seconds) return '0s';
  var h = Math.floor(seconds / 3600);
  var m = Math.floor((seconds % 3600) / 60);
  var s = seconds % 60;
  if (h > 0) return h + 'h ' + m + 'm ' + s + 's';
  if (m > 0) return m + 'm ' + s + 's';
  return s + 's';
}

function renderCharts() {
  if (typeof Chart === 'undefined') {
    console.warn('Chart.js nie jest załadowany. Wykresy niedostępne.');
    return;
  }

  // Zlicz typy principalów
  var principalTypeCounts = { User: 0, SharePointGroup: 0, DomainGroup: 0, SpecialPrincipal: 0, Claim: 0 };
  var permLevelCounts = {};
  var objectTypeCounts = {};
  var uniqueCount = 0, inheritedCount = 0;

  App.flatRows.forEach(function(row) {
    // ObjectType
    var ot = row.obj.ObjectType || 'Unknown';
    objectTypeCounts[ot] = (objectTypeCounts[ot] || 0) + 1;

    // Unique vs Inherited
    if (row.obj.HasUniquePermissions) uniqueCount++;
    else inheritedCount++;

    if (row.assignment) {
      var pt = row.assignment.PrincipalType || 'User';
      principalTypeCounts[pt] = (principalTypeCounts[pt] || 0) + 1;

      if (row.assignment.PermissionLevels) {
        row.assignment.PermissionLevels.forEach(function(pl) {
          if (pl) permLevelCounts[pl] = (permLevelCounts[pl] || 0) + 1;
        });
      }
    }
  });

  var chartColors = ['#0078d4','#7b1fa2','#2e7d32','#e65100','#c2185b','#00695c','#bf360c','#1565c0'];
  var darkChartColors = ['#60b4f0','#ce93d8','#81c784','#ffb74d','#f48fb1','#80cbc4','#ff7043','#64b5f6'];
  var isDark = document.documentElement.getAttribute('data-theme') === 'dark';
  var colors = isDark ? darkChartColors : chartColors;

  var chartDefaults = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        position: 'bottom',
        labels: {
          color: isDark ? '#e5e5e5' : '#323130',
          font: { size: 11 },
          boxWidth: 12,
          padding: 8
        }
      }
    }
  };

  // Wykres 1: Typy principalów
  destroyChart('chartPrincipalTypes');
  var ctx1 = document.getElementById('chartPrincipalTypes');
  if (ctx1) {
    var ptLabels = Object.keys(principalTypeCounts).filter(function(k) { return principalTypeCounts[k] > 0; });
    var ptValues = ptLabels.map(function(k) { return principalTypeCounts[k]; });
    App.charts.principalTypes = new Chart(ctx1, {
      type: 'doughnut',
      data: { labels: ptLabels, datasets: [{ data: ptValues, backgroundColor: colors, borderWidth: 1 }] },
      options: Object.assign({}, chartDefaults)
    });
  }

  // Wykres 2: Poziomy uprawnień (top 8)
  destroyChart('chartPermissionLevels');
  var ctx2 = document.getElementById('chartPermissionLevels');
  if (ctx2) {
    var plEntries = Object.entries(permLevelCounts).sort(function(a,b) { return b[1]-a[1]; }).slice(0, 8);
    var plLabels = plEntries.map(function(e) { return e[0]; });
    var plValues = plEntries.map(function(e) { return e[1]; });
    App.charts.permLevels = new Chart(ctx2, {
      type: 'bar',
      data: {
        labels: plLabels,
        datasets: [{
          label: 'Przypisania',
          data: plValues,
          backgroundColor: colors[0],
          borderRadius: 3
        }]
      },
      options: Object.assign({}, chartDefaults, {
        plugins: Object.assign({}, chartDefaults.plugins, { legend: { display: false } }),
        scales: {
          x: { ticks: { color: isDark ? '#a0a0a0' : '#605e5c', font: { size: 10 } }, grid: { color: isDark ? '#3d3d3d' : '#edebe9' } },
          y: { ticks: { color: isDark ? '#a0a0a0' : '#605e5c', font: { size: 10 } }, grid: { color: isDark ? '#3d3d3d' : '#edebe9' } }
        }
      })
    });
  }

  // Wykres 3: Typy obiektów
  destroyChart('chartObjectTypes');
  var ctx3 = document.getElementById('chartObjectTypes');
  if (ctx3) {
    var otLabels = Object.keys(objectTypeCounts);
    var otValues = otLabels.map(function(k) { return objectTypeCounts[k]; });
    App.charts.objectTypes = new Chart(ctx3, {
      type: 'pie',
      data: { labels: otLabels, datasets: [{ data: otValues, backgroundColor: colors, borderWidth: 1 }] },
      options: Object.assign({}, chartDefaults)
    });
  }

  // Wykres 4: Unique vs Inherited
  destroyChart('chartUniqVsInherited');
  var ctx4 = document.getElementById('chartUniqVsInherited');
  if (ctx4) {
    App.charts.uniqVsInherited = new Chart(ctx4, {
      type: 'doughnut',
      data: {
        labels: ['Unikatowe', 'Dziedziczone'],
        datasets: [{ data: [uniqueCount, inheritedCount], backgroundColor: ['#d84315', '#2e7d32'], borderWidth: 1 }]
      },
      options: Object.assign({}, chartDefaults)
    });
  }
}

function destroyChart(canvasId) {
  var canvas = document.getElementById(canvasId);
  if (canvas) {
    var existing = Chart.getChart(canvas);
    if (existing) existing.destroy();
  }
}

function renderTopUniqueObjects() {
  // Obiekty z największą liczbą principalów
  var counts = {};
  App.data.Objects.forEach(function(obj) {
    if (obj.HasUniquePermissions && obj.Assignments && obj.Assignments.length > 0) {
      counts[obj.ObjectId] = { obj: obj, count: obj.Assignments.length };
    }
  });

  var sorted = Object.values(counts).sort(function(a, b) { return b.count - a.count; }).slice(0, 20);

  var html = '';
  sorted.forEach(function(item, idx) {
    var obj = item.obj;
    var url = obj.FullUrl || obj.ServerRelativeUrl || obj.Title || obj.ObjectId;
    html += '<div class="top-list-item" onclick="showObjectDetail(' + escapeHtml(JSON.stringify(obj.ObjectId)) + ')" title="' + escapeHtml(url) + '">';
    html += '  <div class="top-list-rank">' + (idx + 1) + '</div>';
    html += '  <div class="obj-type-icon obj-icon-' + escapeHtml(getObjectTypeClass(obj.ObjectType)) + '">' + getObjectTypeIcon(obj.ObjectType) + '</div>';
    html += '  <div class="top-list-url">' + escapeHtml(truncateUrl(url, 70)) + '</div>';
    html += '  <div class="top-list-count">' + escapeHtml(String(item.count)) + ' principal(ów)</div>';
    html += '</div>';
  });

  var el = document.getElementById('topUniqueObjects');
  if (el) el.innerHTML = html || '<p class="text-muted">Brak obiektów z unikatowymi uprawnieniami.</p>';
}

function renderAlerts() {
  var alerts = [];
  var stats = App.data.Statistics || {};

  if (stats.ErrorCount > 0) {
    alerts.push({ type: 'danger', icon: '⚠', title: 'Błędy skanowania: ' + stats.ErrorCount, desc: 'Sprawdź plik logu w celu identyfikacji nieskonfigurowanych obiektów.' });
  }

  // Sprawdź osierocone konta
  var orphanedCount = 0;
  App.flatRows.forEach(function(row) {
    if (row.assignment && row.assignment.IsOrphaned) orphanedCount++;
  });
  if (orphanedCount > 0) {
    alerts.push({ type: 'warning', icon: '👤', title: 'Osierocone konta: ' + orphanedCount, desc: 'Wykryto konta, które mogą nie istnieć w Active Directory.' });
  }

  // Sprawdź Full Control poza SC
  var dangerousFC = 0;
  App.flatRows.forEach(function(row) {
    if (row.assignment && row.assignment.PermissionLevels && row.assignment.PermissionLevels.includes('Full Control') &&
        row.obj.ObjectType !== 'SiteCollection' && row.obj.ObjectType !== 'WebApplication' && row.assignment.SourceType === 'Direct') {
      dangerousFC++;
    }
  });
  if (dangerousFC > 0) {
    alerts.push({ type: 'warning', icon: '🔑', title: 'Full Control poza SC: ' + dangerousFC + ' przypisań bezpośrednich', desc: 'Sprawdź czy Full Control na listach/plikach jest zamierzone.' });
  }

  if (alerts.length === 0) {
    alerts.push({ type: 'success', icon: '✓', title: 'Brak alertów krytycznych', desc: 'Skanowanie nie wykryło oczywistych problemów bezpieczeństwa. Przejrzyj raport szczegółowy.' });
  }

  var html = '';
  alerts.forEach(function(a) {
    html += '<div class="alert-item alert-' + escapeHtml(a.type) + '">';
    html += '  <span class="alert-icon">' + a.icon + '</span>';
    html += '  <div><div class="alert-title">' + escapeHtml(a.title) + '</div><div class="alert-desc">' + escapeHtml(a.desc) + '</div></div>';
    html += '</div>';
  });

  var el = document.getElementById('alertsList');
  if (el) el.innerHTML = html;
}

/* ============================================================
   SEKCJA 5: TABELA UPRAWNIEŃ (DataTables)
   ============================================================ */

function initTable(skipPopulate) {
  if (typeof $.fn.DataTable === 'undefined') {
    console.warn('DataTables nie jest załadowane');
    if (!skipPopulate) initTableFallback();
    return;
  }

  // Przy rebuildTable tbody jest już wypełniony przefiltrowanymi wierszami –
  // pomijamy populateTableBody(), żeby nie nadpisać filtrowania danymi z flatRows
  if (!skipPopulate) {
    populateTableBody();
  }

  App.dataTable = $('#permissionsTable').DataTable({
    pageLength: 50,
    lengthMenu: [[25, 50, 100, 500, -1], ['25', '50', '100', '500', 'Wszystkie']],
    ordering: true,
    order: [[1, 'asc']],
    searching: true,
    language: {
      search: 'Szukaj:',
      lengthMenu: 'Pokaż _MENU_ wierszy',
      info: 'Wyświetlono _START_ do _END_ z _TOTAL_ wierszy',
      infoEmpty: 'Brak wyników',
      infoFiltered: '(filtrowanie z _MAX_ wierszy)',
      paginate: { first: '«', previous: '‹', next: '›', last: '»' },
      emptyTable: 'Brak danych spełniających kryteria',
      zeroRecords: 'Nie znaleziono pasujących wyników'
    },
    columnDefs: [
      { orderable: false, targets: [0, 6] },
      { className: 'col-check', targets: [0] }
    ],
    drawCallback: function() {
      updateCheckAllState();
    }
  });

  // Zdarzenia checkboxów – bindujemy tylko przy pierwszej inicjalizacji
  if (!skipPopulate) {
    document.getElementById('checkAll').addEventListener('change', function() {
      var checked = this.checked;
      document.querySelectorAll('#tableBody .row-check').forEach(function(cb) {
        cb.checked = checked;
        var row = cb.closest('tr');
        if (row) row.classList.toggle('selected', checked);
        var groupId = cb.getAttribute('data-group-id');
        if (checked) addGroupToSelection(groupId);
        else removeGroupFromSelection(groupId);
      });
      updateRemediationPanel();
    });

    document.getElementById('btnSelectAll').addEventListener('click', function() {
      document.querySelectorAll('#tableBody .row-check').forEach(function(cb) {
        cb.checked = true;
        var row = cb.closest('tr');
        if (row) row.classList.add('selected');
        addGroupToSelection(cb.getAttribute('data-group-id'));
      });
      updateRemediationPanel();
    });

    document.getElementById('btnDeselectAll').addEventListener('click', function() {
      document.querySelectorAll('#tableBody .row-check').forEach(function(cb) {
        cb.checked = false;
        var row = cb.closest('tr');
        if (row) row.classList.remove('selected');
      });
      App.selectedRows = [];
      updateRemediationPanel();
    });
  }
}

function populateTableBody() {
  var tbody = document.getElementById('tableBody');
  if (!tbody) return;

  var rowMap = {};  // rowId -> flatRow
  App.flatRows.forEach(function(row) { rowMap[row.rowId] = row; });
  App.rowMap = rowMap;

  var html = '';
  App.filteredGroupedRows.forEach(function(gRow) {
    html += buildGroupedTableRow(gRow);
  });
  tbody.innerHTML = html;

  // Event delegation dla checkboxow grup (jeden wiersz = jeden obiekt)
  tbody.addEventListener('change', function(e) {
    if (e.target.classList.contains('row-check')) {
      var groupId = e.target.getAttribute('data-group-id');
      var tr = e.target.closest('tr');
      if (e.target.checked) {
        if (tr) tr.classList.add('selected');
        addGroupToSelection(groupId);
      } else {
        if (tr) tr.classList.remove('selected');
        removeGroupFromSelection(groupId);
      }
      updateRemediationPanel();
    }
  });
}


function buildGroupedTableRow(gRow) {
  var obj = gRow.obj;
  var groupId = gRow.groupId;

  var objectTypeIcon = getObjectTypeIcon(obj.ObjectType) + ' <span class="obj-type-icon">' + escapeHtml(obj.ObjectType || '') + '</span>';

  var locationUrl = getObjectUrl(obj) || obj.SiteCollectionUrl || obj.WebApplicationUrl || '';
  var locationDisplay = truncateUrl(locationUrl, 50);
  var titleDisplay = escapeHtml(obj.Title || obj.Name || obj.FileLeafRef || '-');

  var inheritanceBadge = obj.HasUniquePermissions
    ? '<span class="badge badge-unique">Unikatowe</span>'
    : '<span class="badge badge-inherited">Dziedziczone</span>';

  var assignmentsHtml = '';
  if (!gRow.assignments || gRow.assignments.length === 0) {
    assignmentsHtml = '<span class="text-muted asgn-empty">Brak bezposrednich przypisania (dziedziczenie)</span>';
  } else {
    assignmentsHtml = '<div class="assignments-list">';
    gRow.assignments.forEach(function(aItem) {
      var asgn = aItem.assignment;
      var rowId = aItem.rowId;
      var principalBadge = getPrincipalTypeBadge(asgn.PrincipalType);
      var principalName = asgn.DisplayName || asgn.LoginName || '-';
      var orphanBadge = asgn.IsOrphaned ? '<span class="badge badge-orphaned" title="Konto osierocone">!</span> ' : '';
      var adminBadge = asgn.IsSiteAdmin ? '<span class="badge badge-admin">SC Admin</span> ' : '';
      var levels = (asgn.PermissionLevels || []).join(', ');
      var limitedBadge = asgn.IsLimitedAccessOnly ? ' <span class="badge badge-limited">LA</span>' : '';
      var sourceBadge = getSourceBadge(asgn.SourceType);
      var srcName = asgn.SourceName ? ' <span class="text-muted" style="font-size:11px">' + escapeHtml(truncateUrl(asgn.SourceName, 20)) + '</span>' : '';
      assignmentsHtml += '<div class="assignment-row">'+
        '<span class="asgn-principal">' + principalBadge + ' ' + orphanBadge + adminBadge +
        '<span title="' + escapeHtml(asgn.LoginName || '') + '">' + escapeHtml(truncateUrl(principalName, 28)) + '</span></span>' +
        ' <span class="asgn-sep">&mdash;</span> ' +
        '<span class="asgn-levels"><em>' + escapeHtml(levels) + '</em>' + limitedBadge + '</span>' +
        ' ' + sourceBadge + srcName +
        ' <button class="btn-action btn-action-sm" onclick="addSingleToRemediation(\'' + escapeHtml(rowId) + '\')" title="Dodaj do remediacji">'+
        '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="12" height="12"><path d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z"/></svg>'+
        '</button>'+
        '</div>';
    });
    assignmentsHtml += '</div>';
  }

  var actionsHtml = '<div class="action-cell">'+
    '<button class="btn-action" onclick="showObjectDetail(\'' + escapeHtml(obj.ObjectId) + '\')" title="Szczegoly obiektu">'+
    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>'+
    '</button>'+
    '</div>';

  return '<tr data-group-id="' + escapeHtml(groupId) + '">'+
    '<td class="col-check"><input type="checkbox" class="row-check" data-group-id="' + escapeHtml(groupId) + '"></td>'+
    '<td>' + objectTypeIcon + '</td>'+
    '<td class="cell-url"><a href="' + escapeHtml(locationUrl) + '" title="' + escapeHtml(locationUrl) + '" target="_blank" rel="noopener">' + escapeHtml(locationDisplay) + '</a></td>'+
    '<td>' + titleDisplay + '</td>'+
    '<td>' + inheritanceBadge + '</td>'+
    '<td class="cell-assignments">' + assignmentsHtml + '</td>'+
    '<td>' + actionsHtml + '</td>'+
    '</tr>';
}


/* ============================================================
   SEKCJA 6: FILTRY
   ============================================================ */

function initFilters() {
  // Wypełnij select WebApp
  var fWebApp = document.getElementById('fWebApp');
  App.allWebApps.forEach(function(url) {
    var opt = document.createElement('option');
    opt.value = url;
    opt.textContent = truncateUrl(url, 60);
    if (fWebApp) fWebApp.appendChild(opt);
  });

  // Wypełnij select Site Collection
  var fSite = document.getElementById('fSiteCollection');
  App.allSiteCollections.forEach(function(url) {
    var opt = document.createElement('option');
    opt.value = url;
    opt.textContent = truncateUrl(url, 60);
    if (fSite) fSite.appendChild(opt);
  });

  // Wypełnij select Permission Level
  var fPerm = document.getElementById('fPermLevel');
  App.allPermLevels.forEach(function(pl) {
    var opt = document.createElement('option');
    opt.value = pl;
    opt.textContent = pl;
    if (fPerm) fPerm.appendChild(opt);
  });

  // Filtry Enter
  ['fPrincipalSearch', 'fUrlSearch'].forEach(function(id) {
    var el = document.getElementById(id);
    if (el) el.addEventListener('keydown', function(e) { if (e.key === 'Enter') applyFilters(); });
  });

  document.getElementById('btnApplyFilters').addEventListener('click', applyFilters);
  document.getElementById('btnClearFilters').addEventListener('click', clearFilters);

  // WebApp zmiana -> aktualizuj site collections
  if (fWebApp) fWebApp.addEventListener('change', function() {
    var selected = this.value;
    var fSite = document.getElementById('fSiteCollection');
    if (!fSite) return;
    // Wyczyść opcje poza "Wszystkie"
    while (fSite.options.length > 1) fSite.remove(1);
    // Filtruj SC do wybranej WebApp
    App.allSiteCollections.forEach(function(url) {
      if (!selected || url.indexOf(selected) === 0) {
        var opt = document.createElement('option');
        opt.value = url;
        opt.textContent = truncateUrl(url, 60);
        fSite.appendChild(opt);
      }
    });
  });
}

function applyFilters() {
  var fWebApp = (document.getElementById('fWebApp') || {}).value || '';
  var fSite = (document.getElementById('fSiteCollection') || {}).value || '';
  var fObjType = (document.getElementById('fObjectType') || {}).value || '';
  var fPrincipalType = (document.getElementById('fPrincipalType') || {}).value || '';
  var fInheritance = (document.getElementById('fInheritance') || {}).value || '';
  var fPermLevel = (document.getElementById('fPermLevel') || {}).value || '';
  var fSourceType = (document.getElementById('fSourceType') || {}).value || '';
  var fPrincipalSearch = ((document.getElementById('fPrincipalSearch') || {}).value || '').toLowerCase().trim();
  var fUrlSearch = ((document.getElementById('fUrlSearch') || {}).value || '').toLowerCase().trim();
  var fOnlyUnique = (document.getElementById('fOnlyUnique') || {}).checked;
  var fOnlyDirect = (document.getElementById('fOnlyDirect') || {}).checked;
  var fOnlyLimited = (document.getElementById('fOnlyLimitedAccess') || {}).checked;
  var fOnlyOrphaned = (document.getElementById('fOnlyOrphaned') || {}).checked;

  function objMatchesFilter(obj) {
    if (fWebApp && obj.WebApplicationUrl !== fWebApp) return false;
    if (fSite && obj.SiteCollectionUrl !== fSite) return false;
    if (fObjType && obj.ObjectType !== fObjType) return false;
    if (fInheritance === 'unique' && !obj.HasUniquePermissions) return false;
    if (fInheritance === 'inherited' && obj.HasUniquePermissions) return false;
    if (fOnlyUnique && !obj.HasUniquePermissions) return false;
    if (fUrlSearch) {
      var fullUrl = (obj.FullUrl || obj.ServerRelativeUrl || obj.Title || '').toLowerCase();
      if (fullUrl.indexOf(fUrlSearch) === -1) return false;
    }
    return true;
  }
  function asgnMatchesFilter(asgn) {
    if (!asgn) return false;
    if (fPrincipalType && asgn.PrincipalType !== fPrincipalType) return false;
    if (fSourceType && asgn.SourceType !== fSourceType) return false;
    if (fOnlyDirect && asgn.SourceType !== 'Direct') return false;
    if (fOnlyLimited && !asgn.IsLimitedAccessOnly) return false;
    if (fOnlyOrphaned && !asgn.IsOrphaned) return false;
    if (fPermLevel && !(asgn.PermissionLevels || []).includes(fPermLevel)) return false;
    if (fPrincipalSearch) {
      var login = (asgn.LoginName || '').toLowerCase();
      var display = (asgn.DisplayName || '').toLowerCase();
      var email = (asgn.Email || '').toLowerCase();
      if (login.indexOf(fPrincipalSearch) === -1 && display.indexOf(fPrincipalSearch) === -1 && email.indexOf(fPrincipalSearch) === -1) return false;
    }
    return true;
  }
  var hasAsgnFilter = !!(fPrincipalType || fSourceType || fOnlyDirect || fOnlyLimited || fOnlyOrphaned || fPermLevel || fPrincipalSearch);

  App.filteredRows = App.flatRows.filter(function(row) {
    if (!objMatchesFilter(row.obj)) return false;
    if (hasAsgnFilter) {
      if (!row.assignment) return false;
      return asgnMatchesFilter(row.assignment);
    }
    return true;
  });

  App.filteredGroupedRows = App.groupedRows.filter(function(gRow) {
    if (!objMatchesFilter(gRow.obj)) return false;
    if (hasAsgnFilter) {
      return gRow.assignments.some(function(aItem) { return asgnMatchesFilter(aItem.assignment); });
    }
    return true;
  });

  // Przebuduj  // Przebuduj tabelę DataTables
  rebuildTable();
  showToast('info', 'Filtry zastosowane', 'Grupowanych wyników: ' + App.filteredRows.length);
}

function clearFilters() {
  ['fWebApp','fSiteCollection','fObjectType','fPrincipalType','fInheritance','fPermLevel','fSourceType'].forEach(function(id) {
    var el = document.getElementById(id);
    if (el) el.value = '';
  });
  ['fPrincipalSearch','fUrlSearch'].forEach(function(id) {
    var el = document.getElementById(id);
    if (el) el.value = '';
  });
  ['fOnlyUnique','fOnlyDirect','fOnlyLimitedAccess','fOnlyOrphaned'].forEach(function(id) {
    var el = document.getElementById(id);
    if (el) el.checked = false;
  });

  App.filteredRows = App.flatRows.slice();
  App.filteredGroupedRows = App.groupedRows.slice();
  rebuildTable();
  showToast('info', 'Filtry wyczyszczone', 'Wyświetlono wszystkie wiersze.');
}

function rebuildTable() {
  if (App.dataTable) {
    App.dataTable.destroy();
    App.dataTable = null;
  }
  // Wypełnij tbody przefiltrowanymi wierszami
  var tbody = document.getElementById('tableBody');
  if (tbody) {
    var html = '';
    App.filteredGroupedRows.forEach(function(gRow) { html += buildGroupedTableRow(gRow); });
    tbody.innerHTML = html;
  }
  // Reinicjalizuj tylko DataTables – bez populateTableBody (skipPopulate=true)
  initTable(true);
}

function updateCheckAllState() {
  var checkAll = document.getElementById('checkAll');
  if (!checkAll) return;
  var checks = document.querySelectorAll('#tableBody .row-check');
  var checked = document.querySelectorAll('#tableBody .row-check:checked');
  checkAll.checked = checks.length > 0 && checked.length === checks.length;
  checkAll.indeterminate = checked.length > 0 && checked.length < checks.length;
}

/* ============================================================
   SEKCJA 7: DRZEWO LOKALIZACJI (jsTree)
   ============================================================ */

function initTree() {
  if (typeof $.fn.jstree === 'undefined') {
    console.warn('jsTree nie jest załadowane');
    return;
  }

  var treeData = buildTreeData();

  $('#locationTree').jstree({
    core: {
      data: treeData,
      themes: {
        name: 'default',
        responsive: true
      },
      check_callback: false
    },
    plugins: ['search', 'types'],
    types: {
      webapp:      { icon: 'jstree-folder' },
      site:        { icon: 'jstree-folder' },
      web:         { icon: 'jstree-folder' },
      list:        { icon: 'jstree-file' },
      library:     { icon: 'jstree-folder' },
      folder:      { icon: 'jstree-folder' },
      file:        { icon: 'jstree-file' },
      listitem:    { icon: 'jstree-file' }
    },
    search: { show_only_matches: true }
  }).on('select_node.jstree', function(e, data) {
    // Użyj data.node.data.objectId zamiast manipulacji stringiem na data.node.id
    // Eliminuje ryzyko gdy jsTree zmodyfikuje ID węzła przy renderowaniu
    var objectId = (data.node.data && data.node.data.objectId)
      ? data.node.data.objectId
      : data.node.id.replace('node_', '');
    showTreeNodeDetail('node_' + objectId);
  });

  App.treeInstance = $('#locationTree').jstree(true);

  // Wyszukiwanie w drzewie
  var searchTimeout;
  var treeSearchEl = document.getElementById('treeSearch');
  if (treeSearchEl) {
    treeSearchEl.addEventListener('keyup', function() {
      var val = this.value;
      clearTimeout(searchTimeout);
      searchTimeout = setTimeout(function() {
        $('#locationTree').jstree('search', val);
      }, 250);
    });
  }
}

function normalizeTreeUrl(url) {
  if (!url) return '';
  return url.toLowerCase().replace(/\/$/, '').trim();
}

function buildTreeData() {
  var nodeMap = {};

  // Zbuduj indeks ServerRelativeUrl → ObjectId dla rozwiązywania hierarchii
  var urlIndex = {};
  App.data.Objects.forEach(function(obj) {
    var url = normalizeTreeUrl(obj.ServerRelativeUrl);
    var webApp = (obj.WebApplicationUrl || '').toLowerCase().replace(/\/+$/, '');
    if (url) urlIndex[webApp + '|' + url] = obj.ObjectId;
  });

  var itemTypes = { 'File': true, 'Folder': true, 'ListItem': true, 'Web': true };

  App.data.Objects.forEach(function(obj) {
    var nodeId = 'node_' + obj.ObjectId;
    var parentObjectId = obj.ParentObjectId;

    // Dla plików, folderów i witryn: wyznacz rodzica przez obcięcie ostatniego segmentu URL
    // (ParentObjectId z PS może wskazywać na SC zamiast na rzeczywistego rodzica)
    if (itemTypes[obj.ObjectType] && obj.ServerRelativeUrl) {
      var itemUrl = normalizeTreeUrl(obj.ServerRelativeUrl);
      if (itemUrl) {
        var lastSlash = itemUrl.lastIndexOf('/');
        if (lastSlash > 0) {
          var parentUrl = itemUrl.substring(0, lastSlash);
          var itemWebApp = (obj.WebApplicationUrl || '').toLowerCase().replace(/\/+$/, '');
          var resolved = urlIndex[itemWebApp + '|' + parentUrl];
          if (resolved && resolved !== obj.ObjectId) {
            parentObjectId = resolved;
          }
        }
      }
    }

    var parentNodeId = parentObjectId ? 'node_' + parentObjectId : '#';

    var badge = obj.HasUniquePermissions
      ? ' <span class="tree-badge badge-unique">U</span>'
      : ' <span class="tree-badge badge-inherited">I</span>';

    var assignCount = obj.Assignments ? obj.Assignments.length : 0;
    var countBadge = assignCount > 0 ? ' <span class="tree-badge badge-direct">' + assignCount + '</span>' : '';

    var label = escapeHtml(obj.Title || obj.Name || obj.FileLeafRef || obj.ObjectType || obj.ObjectId);

    nodeMap[obj.ObjectId] = {
      id: nodeId,
      parent: parentNodeId,
      text: label + badge + countBadge,
      type: getJsTreeType(obj.ObjectType),
      data: { objectId: obj.ObjectId }
    };
  });

  // Drugi pass: napraw osierocone węzły (rodzic nie istnieje w drzewie)
  Object.keys(nodeMap).forEach(function(objId) {
    var node = nodeMap[objId];
    if (node.parent === '#') return;
    var parentObjId = node.parent.substring(5); // strip 'node_'
    if (nodeMap[parentObjId]) return; // rodzic istnieje - OK

    // Rodzic nie istnieje - próbuj URL-based fallback
    var obj = App.objectMap[objId];
    var fallback = null;
    if (obj && obj.ServerRelativeUrl) {
      var url = normalizeTreeUrl(obj.ServerRelativeUrl);
      var ls = url.lastIndexOf('/');
      if (ls > 0) {
        var pUrl = url.substring(0, ls);
        var orphanWebApp = (obj.WebApplicationUrl || '').toLowerCase().replace(/\/+$/, '');
        var cand = urlIndex[orphanWebApp + '|' + pUrl];
        if (cand && nodeMap[cand]) fallback = cand;
      }
    }
    node.parent = fallback ? 'node_' + fallback : '#';
  });

  return Object.values(nodeMap);
}

function getJsTreeType(objectType) {
  var map = {
    'WebApplication': 'webapp',
    'SiteCollection': 'site',
    'Web': 'web',
    'List': 'list',
    'Library': 'library',
    'Folder': 'folder',
    'File': 'file',
    'ListItem': 'listitem'
  };
  return map[objectType] || 'web';
}

function treeExpandAll() {
  if (App.treeInstance) App.treeInstance.open_all();
}

function treeCollapseAll() {
  if (App.treeInstance) App.treeInstance.close_all();
}

function showTreeNodeDetail(nodeId) {
  var objectId = nodeId.replace('node_', '');
  var obj = App.data.Objects.find(function(o) { return o.ObjectId === objectId; });

  var panel = document.getElementById('treeDetailPanel');
  if (!panel) { showObjectDetail(objectId); return; }

  if (!obj) {
    panel.innerHTML = '<div class="tree-detail-placeholder"><p>Nie znaleziono obiektu.</p></div>';
    return;
  }

  var url = getObjectUrl(obj);
  var assignCount = obj.Assignments ? obj.Assignments.length : 0;

  var html = '<div style="padding:16px;overflow-y:auto;height:100%;box-sizing:border-box">';

  // Nagłówek
  html += '<div style="margin-bottom:14px;border-bottom:1px solid #edebe9;padding-bottom:12px">';
  html += '<div style="display:flex;align-items:center;gap:8px;margin-bottom:6px">';
  html += getObjectTypeIcon(obj.ObjectType);
  html += '<strong style="font-size:.95rem">' + escapeHtml(obj.Title || obj.Name || obj.FileLeafRef || obj.ObjectId) + '</strong>';
  html += '</div>';
  html += obj.HasUniquePermissions
    ? '<span class="badge badge-unique">Unikatowe ACL</span>'
    : '<span class="badge badge-inherited">Dziedziczone</span>';
  if (url) {
    html += '<div style="margin-top:6px;font-size:11px;color:#605e5c;word-break:break-all">';
    html += '<a href="' + escapeHtml(url) + '" target="_blank" rel="noopener">' + escapeHtml(truncateUrl(url, 55)) + '</a>';
    html += '</div>';
  }
  html += '</div>';

  // Właściwości podstawowe
  var basicProps = [
    ['Typ obiektu', obj.ObjectType],
    ['Web Application', obj.WebApplicationUrl],
    ['Site Collection', obj.SiteCollectionUrl !== obj.WebApplicationUrl ? obj.SiteCollectionUrl : null],
    ['Witryna (Web)', obj.WebUrl !== obj.SiteCollectionUrl ? obj.WebUrl : null],
    ['Lista', obj.ListTitle || null]
  ];
  var propRows = basicProps.filter(function(p) { return p[1]; });
  if (propRows.length > 0) {
    html += '<table style="width:100%;border-collapse:collapse;font-size:12px;margin-bottom:14px">';
    propRows.forEach(function(p) {
      html += '<tr><td style="padding:3px 6px 3px 0;color:#605e5c;white-space:nowrap;vertical-align:top">' + escapeHtml(p[0]) + '</td>';
      html += '<td style="padding:3px 0;word-break:break-all;font-size:11px">' + escapeHtml(p[1]) + '</td></tr>';
    });
    html += '</table>';
  }

  // Przypisania
  if (assignCount > 0) {
    html += '<div style="font-weight:600;font-size:13px;margin-bottom:6px">Przypisania (' + assignCount + ')</div>';
    html += '<table class="detail-table" style="font-size:12px;width:100%">';
    html += '<thead><tr><th>Principal</th><th>Poziomy uprawnień</th><th>Źródło</th></tr></thead><tbody>';
    obj.Assignments.forEach(function(a) {
      html += '<tr>';
      html += '<td style="vertical-align:top">';
      html += '<strong>' + escapeHtml(a.DisplayName || a.LoginName || '-') + '</strong>';
      if (a.LoginName && a.LoginName !== (a.DisplayName || '')) {
        html += '<br><code style="font-size:10px;color:#605e5c">' + escapeHtml(a.LoginName) + '</code>';
      }
      if (a.IsOrphaned) html += ' <span class="badge badge-orphaned" title="Osierocone">!</span>';
      if (a.IsSiteAdmin) html += ' <span class="badge badge-admin">SC Admin</span>';
      html += '<br>' + getPrincipalTypeBadge(a.PrincipalType);
      html += '</td>';
      html += '<td>' + escapeHtml((a.PermissionLevels || []).join(', ') || '-');
      if (a.IsLimitedAccessOnly) html += ' <span class="badge badge-limited">LA</span>';
      html += '</td>';
      html += '<td>' + getSourceBadge(a.SourceType);
      if (a.SourceName) html += '<br><small style="color:#605e5c">' + escapeHtml(truncateUrl(a.SourceName, 28)) + '</small>';
      html += '</td>';
      html += '</tr>';
    });
    html += '</tbody></table>';
  } else if (!obj.HasUniquePermissions) {
    html += '<div style="color:#605e5c;font-size:13px;padding:6px 0">';
    html += 'Dziedziczy uprawnienia z:<br>';
    html += '<a href="' + escapeHtml(obj.InheritsFromUrl || '#') + '" target="_blank" rel="noopener" style="font-size:11px;word-break:break-all">';
    html += escapeHtml(obj.InheritsFromUrl || 'obiektu nadrzędnego') + '</a></div>';
  } else {
    html += '<div style="color:#605e5c;font-size:13px">Brak bezpośrednich przypisań.</div>';
  }

  html += buildObjectContentSection(obj, true);

  // Przycisk pełnych szczegółów
  html += '<div style="margin-top:14px;border-top:1px solid #edebe9;padding-top:10px">';
  html += '<button class="btn btn-sm" onclick="showObjectDetail(\'' + escapeHtml(objectId) + '\')">Pełne szczegóły (modal)...</button>';
  html += '</div>';

  html += '</div>';
  panel.innerHTML = html;
}

/* ============================================================
   SEKCJA 8: WYSZUKIWANIE UŻYTKOWNIKA
   ============================================================ */

function initUserLookup() {
  var input = document.getElementById('userLookupInput');
  var btn = document.getElementById('btnUserLookup');
  var suggestionsEl = document.getElementById('lookupSuggestions');

  if (btn) btn.addEventListener('click', performUserLookup);
  if (input) {
    input.addEventListener('keydown', function(e) {
      if (e.key === 'Enter') performUserLookup();
    });

    // Autocomplete
    var autocompleteTimeout;
    input.addEventListener('input', function() {
      var val = this.value.trim().toLowerCase();
      clearTimeout(autocompleteTimeout);
      if (val.length < 2) {
        if (suggestionsEl) suggestionsEl.style.display = 'none';
        return;
      }
      autocompleteTimeout = setTimeout(function() {
        showAutocompleteSuggestions(val);
      }, 200);
    });
  }

  var exportBtn = document.getElementById('btnExportLookupCsv');
  if (exportBtn) exportBtn.addEventListener('click', exportLookupResultsCsv);
}

function showAutocompleteSuggestions(query) {
  var suggestionsEl = document.getElementById('lookupSuggestions');
  if (!suggestionsEl) return;

  var matches = App.allPrincipals.filter(function(p) {
    return (p.loginName && p.loginName.toLowerCase().indexOf(query) !== -1) ||
           (p.displayName && p.displayName.toLowerCase().indexOf(query) !== -1) ||
           (p.email && p.email.toLowerCase().indexOf(query) !== -1);
  }).slice(0, 10);

  if (matches.length === 0) {
    suggestionsEl.style.display = 'none';
    return;
  }

  var html = '';
  matches.forEach(function(p) {
    var ptIcon = getPrincipalTypeIconStr(p.principalType);
    html += '<div class="lookup-suggestion-item" onclick="selectSuggestion(' + escapeHtml(JSON.stringify(p.loginName)) + ')">';
    html += '  <span>' + ptIcon + '</span>';
    html += '  <div>';
    html += '    <div class="lookup-suggestion-name">' + escapeHtml(p.displayName || p.loginName) + '</div>';
    html += '    <div class="lookup-suggestion-login">' + escapeHtml(p.loginName) + '</div>';
    html += '  </div>';
    html += '</div>';
  });

  suggestionsEl.innerHTML = html;
  suggestionsEl.style.display = '';
}

function selectSuggestion(loginName) {
  var input = document.getElementById('userLookupInput');
  var suggestionsEl = document.getElementById('lookupSuggestions');
  if (input) input.value = loginName;
  if (suggestionsEl) suggestionsEl.style.display = 'none';
  performUserLookup();
}

function performUserLookup() {
  var input = document.getElementById('userLookupInput');
  var suggestionsEl = document.getElementById('lookupSuggestions');
  if (suggestionsEl) suggestionsEl.style.display = 'none';

  var query = (input ? input.value.trim().toLowerCase() : '').trim();
  if (!query) {
    showToast('warning', 'Brak zapytania', 'Wpisz nazwę użytkownika lub grupy.');
    return;
  }

  // Znajdź wszystkie wiersze dla danego użytkownika/grupy
  var matches = App.flatRows.filter(function(row) {
    if (!row.assignment) return false;
    var a = row.assignment;
    return (a.LoginName && a.LoginName.toLowerCase().indexOf(query) !== -1) ||
           (a.DisplayName && a.DisplayName.toLowerCase().indexOf(query) !== -1) ||
           (a.Email && a.Email.toLowerCase().indexOf(query) !== -1);
  });

  renderLookupResults(query, matches);

  var exportBtn = document.getElementById('btnExportLookupCsv');
  if (exportBtn) exportBtn.style.display = matches.length > 0 ? '' : 'none';
}

function renderLookupResults(query, matches) {
  var resultsEl = document.getElementById('lookupResults');
  if (!resultsEl) return;

  if (matches.length === 0) {
    resultsEl.innerHTML = '<div class="lookup-placeholder"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg><p>Nie znaleziono wyników dla: <strong>' + escapeHtml(query) + '</strong></p></div>';
    return;
  }

  // Pobierz dane principala z pierwszego wyniku
  var firstAsgn = matches[0].assignment;
  var principalName = firstAsgn.DisplayName || firstAsgn.LoginName || query;
  var principalLogin = firstAsgn.LoginName || '';
  var principalType = firstAsgn.PrincipalType || '';
  var principalEmail = firstAsgn.Email || '';

  var html = '';
  html += '<div class="lookup-result-header">';
  html += '  ' + getPrincipalTypeIconStr(principalType);
  html += '  <div class="lookup-result-name">' + escapeHtml(principalName) + '</div>';
  html += '  <div class="lookup-result-login">' + escapeHtml(principalLogin) + (principalEmail ? ' &bull; ' + escapeHtml(principalEmail) : '') + '</div>';
  html += '  <div>' + getPrincipalTypeBadge(principalType) + '</div>';
  html += '</div>';

  html += '<p class="info-text">Znaleziono dostęp do <strong>' + escapeHtml(String(matches.length)) + '</strong> obiektów.</p>';

  html += '<table class="lookup-access-table">';
  html += '<thead><tr><th>Typ obiektu</th><th>Lokalizacja</th><th>Tytuł</th><th>Uprawnienia</th><th>Źródło</th><th>Przez</th></tr></thead>';
  html += '<tbody>';

  matches.forEach(function(row) {
    var obj = row.obj;
    var asgn = row.assignment;
    var url = getObjectUrl(obj);
    var sourceExplanation = explainAccessSource(asgn);

    html += '<tr>';
    html += '  <td>' + getObjectTypeIcon(obj.ObjectType) + ' ' + escapeHtml(obj.ObjectType || '') + '</td>';
    html += '  <td><a href="' + escapeHtml(url) + '" target="_blank" title="' + escapeHtml(url) + '" rel="noopener">' + escapeHtml(truncateUrl(url, 50)) + '</a></td>';
    html += '  <td>' + escapeHtml(obj.Title || obj.Name || '-') + '</td>';
    html += '  <td>' + escapeHtml((asgn.PermissionLevels || []).join(', ') || '-') + (asgn.IsLimitedAccessOnly ? ' <span class="badge badge-limited">LA</span>' : '') + '</td>';
    html += '  <td>' + getSourceBadge(asgn.SourceType) + '</td>';
    html += '  <td>' + escapeHtml(sourceExplanation) + '</td>';
    html += '</tr>';
  });

  html += '</tbody></table>';

  resultsEl.innerHTML = html;
  App.lastLookupResults = matches;
}

function explainAccessSource(asgn) {
  if (!asgn) return '';
  switch (asgn.SourceType) {
    case 'Direct': return 'Bezpośrednie przypisanie';
    case 'ViaSharePointGroup': return 'Przez grupę SharePoint: ' + (asgn.SourceName || '');
    case 'ViaDomainGroup': return 'Przez grupę AD: ' + (asgn.SourceName || '');
    case 'Inherited': return 'Odziedziczone z: ' + (asgn.SourceName || '');
    default: return asgn.SourceType || '-';
  }
}

function exportLookupResultsCsv() {
  var rows = App.lastLookupResults || [];
  if (rows.length === 0) return;

  var headers = ['ObjectType','WebApplicationUrl','SiteCollectionUrl','WebUrl','FullUrl','Title','HasUniquePermissions','PrincipalType','LoginName','DisplayName','Email','SourceType','SourceName','PermissionLevels','IsLimitedAccessOnly','IsOrphaned'];
  var csvRows = [headers.join(';')];

  rows.forEach(function(row) {
    var obj = row.obj;
    var a = row.assignment;
    var cells = [
      obj.ObjectType, obj.WebApplicationUrl, obj.SiteCollectionUrl, obj.WebUrl, obj.FullUrl, obj.Title,
      obj.HasUniquePermissions,
      a ? a.PrincipalType : '', a ? a.LoginName : '', a ? a.DisplayName : '', a ? a.Email : '',
      a ? a.SourceType : '', a ? a.SourceName : '',
      a ? (a.PermissionLevels || []).join('|') : '',
      a ? a.IsLimitedAccessOnly : '', a ? a.IsOrphaned : ''
    ];
    csvRows.push(cells.map(function(c) { return '"' + String(c || '').replace(/"/g, '""') + '"'; }).join(';'));
  });

  downloadText(csvRows.join('\r\n'), 'UserLookup_' + new Date().toISOString().slice(0,10) + '.csv', 'text/csv;charset=utf-8;');
}

/* ============================================================
   SEKCJA 9: MODAL SZCZEGÓŁÓW OBIEKTU
   ============================================================ */

function showObjectDetail(objectId) {
  var obj = App.data.Objects.find(function(o) { return o.ObjectId === objectId; });
  if (!obj) {
    showToast('warning', 'Nie znaleziono', 'Obiekt o ID ' + objectId + ' nie istnieje w danych.');
    return;
  }

  App.detailModalObject = obj;

  var modal = document.getElementById('detailModal');
  var modalTitle = document.getElementById('modalTitle');
  var modalBody = document.getElementById('modalBody');

  if (modalTitle) setTextSafe(modalTitle, (obj.ObjectType || '') + ': ' + (obj.Title || obj.Name || obj.ObjectId));

  // Breadcrumb
  var breadcrumb = buildBreadcrumb(obj);

  var html = '';
  html += '<div class="breadcrumb">' + breadcrumb + '</div>';

  // Tabela właściwości
  html += '<div class="detail-section-title">Właściwości obiektu</div>';
  html += '<table class="detail-table">';

  var props = [
    ['ID obiektu', obj.ObjectId],
    ['Typ', obj.ObjectType],
    ['Tytuł', obj.Title],
    ['Nazwa', obj.Name],
    ['Pełny URL', obj.FullUrl],
    ['URL do otwarcia', getObjectUrl(obj)],
    ['Server Relative URL', obj.ServerRelativeUrl],
    ['Web Application', obj.WebApplicationUrl],
    ['Site Collection', obj.SiteCollectionUrl],
    ['Witryna (Web)', obj.WebUrl],
    ['Lista', obj.ListTitle + (obj.ListId ? ' (' + obj.ListId + ')' : '')],
    ['Element ID', obj.ItemId],
    ['Plik/LeafRef', obj.FileLeafRef],
    ['Ukryty', obj.IsHidden ? 'Tak' : 'Nie'],
    ['Systemowy', obj.IsSystem ? 'Tak' : 'Nie'],
    ['Katalog', obj.IsCatalog ? 'Tak' : 'Nie'],
    ['Site Assets', obj.IsSiteAssets ? 'Tak' : 'Nie'],
    ['Unikatowe uprawnienia', obj.HasUniquePermissions ? '✓ TAK' : '✗ Nie (dziedziczy)'],
    ['Dziedziczy z URL', obj.InheritsFromUrl || '-'],
    ['Pierwszy unikatowy przodek', obj.FirstUniqueAncestorUrl || '-'],
    ['Timestamp skanu', obj.ScanTimestamp]
  ];

  props.forEach(function(p) {
    if (p[1] === undefined || p[1] === null || p[1] === '') return;
    html += '<tr><th>' + escapeHtml(p[0]) + '</th><td>' + escapeHtml(String(p[1])) + '</td></tr>';
  });

  html += '</table>';

  // Przypisania uprawnień
  if (obj.Assignments && obj.Assignments.length > 0) {
    html += '<div class="detail-section-title">Przypisania uprawnień (' + obj.Assignments.length + ')</div>';
    html += '<table class="detail-table">';
    html += '<thead><tr><th>Typ</th><th>Principal</th><th>Poziomy</th><th>Źródło</th><th>Ścieżka dziedziczenia</th></tr></thead>';
    html += '<tbody>';
    obj.Assignments.forEach(function(a) {
      html += '<tr>';
      html += '<td>' + getPrincipalTypeBadge(a.PrincipalType) + '</td>';
      html += '<td>';
      html += '<strong>' + escapeHtml(a.DisplayName || a.LoginName || '-') + '</strong>';
      if (a.LoginName) html += '<br><code>' + escapeHtml(a.LoginName) + '</code>';
      if (a.Email) html += '<br>' + escapeHtml(a.Email);
      if (a.IsSiteAdmin) html += ' <span class="badge badge-admin">SC Admin</span>';
      if (a.IsOrphaned) html += ' <span class="badge badge-orphaned">Orphaned</span>';
      if (a.IsUnresolved) html += ' <span class="badge badge-orphaned">Unresolved</span>';
      html += '</td>';
      html += '<td>' + escapeHtml((a.PermissionLevels || []).join(', ')) + (a.IsLimitedAccessOnly ? ' <span class="badge badge-limited">LA</span>' : '') + '</td>';
      html += '<td>' + getSourceBadge(a.SourceType) + ' ' + escapeHtml(a.SourceName || '') + '</td>';
      html += '<td><code>' + escapeHtml((a.InheritancePath || []).join(' > ') || '-') + '</code></td>';
      html += '</tr>';
    });
    html += '</tbody></table>';
  } else if (!obj.HasUniquePermissions) {
    html += '<div class="detail-section-title">Dziedziczenie</div>';
    html += '<p class="info-text">Ten obiekt dziedziczy uprawnienia z: <a href="' + escapeHtml(obj.InheritsFromUrl || '#') + '" target="_blank">' + escapeHtml(obj.InheritsFromUrl || 'obiektu nadrzędnego') + '</a></p>';
  } else {
    html += '<p class="info-text text-muted">Brak przypisań uprawnień na tym obiekcie.</p>';
  }

  html += buildObjectContentSection(obj, false);

  if (modalBody) modalBody.innerHTML = html;
  if (modal) modal.style.display = 'flex';
}

function buildBreadcrumb(obj) {
  var parts = [];
  if (obj.WebApplicationUrl) parts.push({ label: 'Web App', value: obj.WebApplicationUrl });
  if (obj.SiteCollectionUrl && obj.SiteCollectionUrl !== obj.WebApplicationUrl) parts.push({ label: 'SC', value: obj.SiteCollectionUrl });
  if (obj.WebUrl && obj.WebUrl !== obj.SiteCollectionUrl) parts.push({ label: 'Web', value: obj.WebUrl });
  if (obj.ListTitle) parts.push({ label: 'Lista', value: obj.ListTitle });
  if (obj.Title && obj.Title !== obj.ListTitle) parts.push({ label: obj.ObjectType, value: obj.Title });

  return parts.map(function(p, i) {
    var isLast = i === parts.length - 1;
    return '<span class="breadcrumb-item' + (isLast ? ' active' : '') + '" title="' + escapeHtml(p.value) + '">' + escapeHtml(p.label + ': ' + truncateUrl(p.value, 30)) + '</span>'
      + (isLast ? '' : '<span class="breadcrumb-sep">›</span>');
  }).join('');
}

function closeDetailModal() {
  var modal = document.getElementById('detailModal');
  if (modal) modal.style.display = 'none';
  App.detailModalObject = null;
}

function exportObjectDetailJson() {
  if (!App.detailModalObject) return;
  downloadText(JSON.stringify(App.detailModalObject, null, 2), 'Object_' + App.detailModalObject.ObjectId + '.json', 'application/json');
}

// Zamknij modal klawiszem Escape
document.addEventListener('keydown', function(e) {
  if (e.key === 'Escape') closeDetailModal();
});

/* ============================================================
   SEKCJA 10: REMEDIACJA
   ============================================================ */

function initRemediationTab() {
  var dryRunToggle = document.getElementById('dryRunToggle');
  if (dryRunToggle) {
    dryRunToggle.addEventListener('change', updateModeIndicator);
  }

  updateModeIndicator();
  updateRemediationPanel();
  detectSharePointMode();
  autoConnectDetectedSharePoint();
}

function updateModeIndicator() {
  var isDryRun = (document.getElementById('dryRunToggle') || {}).checked !== false;
  var indicator = document.getElementById('modeIndicator');
  if (!indicator) return;

  if (isDryRun) {
    indicator.innerHTML = '<div class="mode-dryrun"><span class="mode-icon">🔍</span> <div><strong>Tryb DRY-RUN aktywny</strong><br><small>Skrypt nie wprowadzi żadnych zmian. Wymagane ponowne uruchomienie z parametrem -DryRun $false</small></div></div>';
  } else {
    indicator.innerHTML = '<div class="mode-live"><span class="mode-icon">⚠</span> <div><strong>TRYB LIVE - rzeczywiste zmiany!</strong><br><small>Wygenerowany skrypt wprowadzi zmiany po uruchomieniu. Sprawdź plan przed wykonaniem.</small></div></div>';
  }
}

function addToRemediationFromRow(rowId) {
  addSingleToRemediation(rowId);
}

function addSingleToRemediation(rowId) {
  if (!App.rowMap || !App.rowMap[rowId]) return;
  var flatRow = App.rowMap[rowId];
  if (!App.selectedRows.some(function(r) { return r.rowId === rowId; })) {
    App.selectedRows.push(flatRow);
  }
  var groupId = flatRow.obj ? flatRow.obj.ObjectId : null;
  if (groupId) {
    var groupCb = document.querySelector('.row-check[data-group-id="' + CSS.escape(groupId) + '"]');
    if (groupCb) {
      groupCb.checked = true;
      var tr = groupCb.closest('tr');
      if (tr) tr.classList.add('selected');
    }
  }
  updateRemediationPanel();
  showToast('info', 'Dodano do remediacji', 'Przejdz do zakladki Remediacja.');
}

function addGroupToSelection(groupId) {
  if (!App.groupedRowMap || !App.groupedRowMap[groupId]) return;
  var gRow = App.groupedRowMap[groupId];
  if (gRow.assignments.length > 0) {
    gRow.assignments.forEach(function(aItem) {
      if (!App.selectedRows.some(function(r) { return r.rowId === aItem.rowId; })) {
        App.selectedRows.push({ obj: gRow.obj, assignment: aItem.assignment, rowId: aItem.rowId });
      }
    });
  } else {
    var noAsgnRowId = groupId + '_empty';
    if (!App.selectedRows.some(function(r) { return r.rowId === noAsgnRowId; })) {
      App.selectedRows.push({ obj: gRow.obj, assignment: null, rowId: noAsgnRowId });
    }
  }
}

function removeGroupFromSelection(groupId) {
  if (!App.groupedRowMap || !App.groupedRowMap[groupId]) return;
  var gRow = App.groupedRowMap[groupId];
  var asgnRowIds = {};
  gRow.assignments.forEach(function(a) { asgnRowIds[a.rowId] = true; });
  asgnRowIds[groupId + '_empty'] = true;
  App.selectedRows = App.selectedRows.filter(function(r) { return !asgnRowIds[r.rowId]; });
}


function addToSelection(rowId) {
  if (!App.rowMap || !App.rowMap[rowId]) return;
  if (!App.selectedRows.some(function(r) { return r.rowId === rowId; })) {
    App.selectedRows.push(App.rowMap[rowId]);
  }
}

function removeFromSelection(rowId) {
  App.selectedRows = App.selectedRows.filter(function(r) { return r.rowId !== rowId; });
}

function clearRemediationSelection() {
  App.selectedRows = [];
  document.querySelectorAll('.row-check').forEach(function(cb) {
    cb.checked = false;
    var row = cb.closest('tr');
    if (row) row.classList.remove('selected');
  });
  updateRemediationPanel();
}

function updateRemediationPanel() {
  var count = App.selectedRows.length;
  var emptyEl = document.getElementById('selectedItemsPanel');
  var listEl = document.getElementById('selectedItemsList');
  var container = document.getElementById('selectedItemsContainer');
  var countEl = document.getElementById('selectedCount');

  if (count === 0) {
    if (emptyEl) emptyEl.style.display = '';
    if (listEl) listEl.style.display = 'none';
    return;
  }

  if (emptyEl) emptyEl.style.display = 'none';
  if (listEl) listEl.style.display = '';
  if (countEl) setTextSafe(countEl, count + ' element' + (count === 1 ? '' : 'ów') + ' wybranych');

  if (!container) return;

  var html = '';
  App.selectedRows.forEach(function(row) {
    var obj = row.obj;
    var asgn = row.assignment;
    var url = getObjectUrl(obj);
    var principal = asgn ? (asgn.DisplayName || asgn.LoginName || '-') : 'Brak przypisań';

    html += '<div class="selected-item-row">';
    html += '  <div class="selected-item-info">';
    html += '    <span class="selected-item-url" title="' + escapeHtml(url) + '">' + escapeHtml(truncateUrl(url, 55)) + '</span>';
    html += '    <span class="selected-item-principal">' + getObjectTypeIcon(obj.ObjectType) + ' ' + escapeHtml(obj.ObjectType || '') + ' &bull; ' + escapeHtml(principal) + '</span>';
    html += '  </div>';
    html += '  <button class="selected-item-remove" onclick="removeFromSelectionAndUpdate(\'' + escapeHtml(row.rowId) + '\')" title="Usuń z wyboru">&times;</button>';
    html += '</div>';
  });

  container.innerHTML = html;
}

function removeFromSelectionAndUpdate(rowId) {
  removeFromSelection(rowId);
  // Check if any assignments for the same object remain; if not, uncheck group row
  var flatRow = App.rowMap ? App.rowMap[rowId] : null;
  if (flatRow && flatRow.obj) {
    var groupId = flatRow.obj.ObjectId;
    var anyRemain = App.selectedRows.some(function(r) { return r.obj && r.obj.ObjectId === groupId; });
    if (!anyRemain) {
      var groupCb = document.querySelector('.row-check[data-group-id="' + CSS.escape(groupId) + '"]');
      if (groupCb) {
        groupCb.checked = false;
        var groupTr = groupCb.closest('tr');
        if (groupTr) groupTr.classList.remove('selected');
      }
    }
  }
  var cb = document.querySelector('.row-check[data-group-id="' + CSS.escape(rowId) + '"]');
  if (cb) {
    cb.checked = false;
    var tr = cb.closest('tr');
    if (tr) tr.classList.remove('selected');
  }
  updateRemediationPanel();
}

function generateRemediationScript() {
  if (App.selectedRows.length === 0) {
    showToast('warning', 'Brak wybranych elementów', 'Zaznacz elementy w tabeli uprawnień przed generowaniem skryptu.');
    return;
  }

  var action = (document.getElementById('remediationAction') || {}).value;
  if (!action) {
    showToast('warning', 'Wybierz akcję', 'Wybierz akcję remediacyjną z listy.');
    return;
  }

  var reason = (document.getElementById('remediationReason') || {}).value || 'Remediacja wygenerowana przez SharePoint Permission Analyzer';
  var isDryRun = (document.getElementById('dryRunToggle') || {}).checked !== false;

  // Walidacja bezpieczeństwa - ochrona przed usunięciem chronionych kont
  var protectedPatterns = [
    /^SHAREPOINT\\system$/i,
    /^NT AUTHORITY\\/i,
    /^SHAREPOINT\\/i
  ];

  var validatedPlan = [];
  var skippedProtected = [];

  App.selectedRows.forEach(function(row) {
    var obj = row.obj;
    var asgn = row.assignment;

    if (!asgn && action !== 'RestoreInheritance') {
      return; // Pomiń wiersze bez przypisań dla akcji innych niż RestoreInheritance
    }

    var loginName = asgn ? (asgn.LoginName || '') : '';
    var isProtected = protectedPatterns.some(function(p) { return p.test(loginName); });

    if (isProtected) {
      skippedProtected.push(loginName);
      return;
    }

    validatedPlan.push({
      Action: action,
      ObjectId: obj.ObjectId,
      ObjectType: obj.ObjectType,
      WebApplicationUrl: obj.WebApplicationUrl || '',
      SiteCollectionUrl: obj.SiteCollectionUrl || '',
      WebUrl: obj.WebUrl || '',
      FullUrl: obj.FullUrl || obj.ServerRelativeUrl || '',
      ServerRelativeUrl: obj.ServerRelativeUrl || '',
      ListId: obj.ListId || '',
      ItemId: obj.ItemId || '',
      PrincipalLoginName: loginName,
      PrincipalDisplayName: asgn ? (asgn.DisplayName || '') : '',
      SharePointGroupName: (asgn && asgn.PrincipalType === 'SharePointGroup') ? (asgn.DisplayName || '') : '',
      PermissionLevels: asgn ? (asgn.PermissionLevels || []) : [],
      Reason: reason,
      GeneratedAt: new Date().toISOString()
    });
  });

  if (skippedProtected.length > 0) {
    showToast('warning', 'Pominięto ' + skippedProtected.length + ' chronionych kont', skippedProtected.join(', '));
  }

  if (validatedPlan.length === 0) {
    showToast('warning', 'Brak operacji do wykonania', 'Wszystkie wybrane elementy zostały pominięte (chronione konta lub brak przypisań).');
    return;
  }

  // Pokaż podgląd
  showRemediationPreview(validatedPlan, action, isDryRun);

  // Wygeneruj skrypt PS1
  var script = buildPs1Script(validatedPlan, isDryRun, reason);
  var filename = 'Remediation_' + action + '_' + new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-') + '.ps1';
  downloadText(script, filename, 'text/plain;charset=utf-8;');

  showToast('success', 'Skrypt wygenerowany', 'Pobierz plik ' + filename + ' i uruchom go na serwerze SharePoint jako Farm Administrator.');
}

function showRemediationPreview(plan, action, isDryRun) {
  var previewEl = document.getElementById('remediationPreview');
  var contentEl = document.getElementById('previewContent');
  if (!previewEl || !contentEl) return;

  var html = '';
  html += '<p class="info-text"><strong>Akcja:</strong> ' + escapeHtml(action) + ' &bull; <strong>Tryb:</strong> ' + (isDryRun ? '🔍 DRY-RUN' : '⚠ LIVE') + ' &bull; <strong>Operacji:</strong> ' + plan.length + '</p>';

  plan.slice(0, 10).forEach(function(op) {
    html += '<div class="preview-item">' + escapeHtml(op.Action) + ' | ' + escapeHtml(truncateUrl(op.FullUrl, 50)) + ' | ' + escapeHtml(op.PrincipalLoginName || 'N/A') + '</div>';
  });

  if (plan.length > 10) {
    html += '<div class="preview-item text-muted">... i ' + (plan.length - 10) + ' więcej operacji</div>';
  }

  contentEl.innerHTML = html;
  previewEl.style.display = '';
}

function buildPs1Script(plan, isDryRun, reason) {
  var now = new Date().toISOString();
  var planJson = JSON.stringify(plan, null, 2);
  var dryRunStr = isDryRun ? '$true' : '$false';

  return '#Requires -Version 5.1\r\n'
    + '<#\r\n'
    + '.SYNOPSIS\r\n'
    + '    SharePoint Permission Remediation Script\r\n'
    + '.DESCRIPTION\r\n'
    + '    Wygenerowany przez SharePoint Permission Analyzer\r\n'
    + '    Data: ' + now + '\r\n'
    + '    Powod: ' + reason + '\r\n'
    + '    Operacji: ' + plan.length + '\r\n'
    + '    Tryb domyslny: ' + (isDryRun ? 'DRY-RUN' : 'LIVE') + '\r\n'
    + '\r\n'
    + '    INSTRUKCJA URUCHOMIENIA:\r\n'
    + '    1. Uruchom na serwerze SharePoint jako Farm Administrator\r\n'
    + '    2. NAJPIERW przetestuj z -DryRun $true\r\n'
    + '    3. Po weryfikacji uruchom z -DryRun $false\r\n'
    + '    4. Sprawdz transcript log po zakonczeniu\r\n'
    + '\r\n'
    + '    PRZYKLAD:\r\n'
    + '      .\\' + 'Remediation.ps1 -DryRun $true   # Symulacja\r\n'
    + '      .\\' + 'Remediation.ps1 -DryRun $false  # Rzeczywiste zmiany\r\n'
    + '#>\r\n'
    + '\r\n'
    + '[CmdletBinding(SupportsShouldProcess = $true)]\r\n'
    + 'param(\r\n'
    + '    [bool]$DryRun = ' + dryRunStr + ',\r\n'
    + '    [string]$TranscriptPath = ""\r\n'
    + ')\r\n'
    + '\r\n'
    + 'Set-StrictMode -Version Latest\r\n'
    + '$ErrorActionPreference = "Stop"\r\n'
    + '\r\n'
    + '# Transcript\r\n'
    + 'if (-not $TranscriptPath) {\r\n'
    + '    $TranscriptPath = Join-Path $PSScriptRoot ("Remediation_" + (Get-Date -Format "yyyy-MM-dd_HH-mm-ss") + ".log")\r\n'
    + '}\r\n'
    + 'Start-Transcript -Path $TranscriptPath -Force | Out-Null\r\n'
    + 'Write-Host "Transcript: $TranscriptPath"\r\n'
    + 'Write-Host "Tryb: $(if ($DryRun) { "DRY-RUN" } else { "LIVE" })"\r\n'
    + 'Write-Host "Operacji: ' + plan.length + '"\r\n'
    + '\r\n'
    + '# Zaladuj SharePoint\r\n'
    + 'if (-not (Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue)) {\r\n'
    + '    Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop\r\n'
    + '}\r\n'
    + '\r\n'
    + '# Whitelist chronionych kont\r\n'
    + '$ProtectedPatterns = @("^SHAREPOINT\\\\\\\\system$", "^NT AUTHORITY\\\\\\\\", "^SHAREPOINT\\\\\\\\")\r\n'
    + '\r\n'
    + 'function Test-IsProtected { param([string]$Login)\r\n'
    + '    foreach ($p in $ProtectedPatterns) { if ($Login -match $p) { return $true } }\r\n'
    + '    return $false\r\n'
    + '}\r\n'
    + '\r\n'
    + '# Plan remediacji (wygenerowany automatycznie)\r\n'
    + '$Plan = @\'\r\n'
    + planJson + '\r\n'
    + '\'@ | ConvertFrom-Json\r\n'
    + '\r\n'
    + '$Success = 0; $Errors = 0; $Skipped = 0\r\n'
    + '\r\n'
    + 'foreach ($Op in $Plan) {\r\n'
    + '    Write-Host ""\r\n'
    + '    Write-Host "Operacja: [$($Op.Action)] $($Op.FullUrl)" -ForegroundColor Cyan\r\n'
    + '    Write-Host "  Principal: $($Op.PrincipalLoginName)"\r\n'
    + '\r\n'
    + '    # Sprawdz ochrone\r\n'
    + '    if ($Op.PrincipalLoginName -and (Test-IsProtected -Login $Op.PrincipalLoginName)) {\r\n'
    + '        Write-Host "  [POMINIETO] Chronione konto: $($Op.PrincipalLoginName)" -ForegroundColor Yellow\r\n'
    + '        $Skipped++; continue\r\n'
    + '    }\r\n'
    + '\r\n'
    + '    try {\r\n'
    + '        if ($DryRun) {\r\n'
    + '            Write-Host "  [DRY-RUN] Wykonalby: $($Op.Action) na $($Op.FullUrl)" -ForegroundColor Yellow\r\n'
    + '            $Success++\r\n'
    + '        } else {\r\n'
    + '            $SPSite = $null; $SPWeb = $null\r\n'
    + '            try {\r\n'
    + '                $SPSite = New-Object Microsoft.SharePoint.SPSite($Op.SiteCollectionUrl)\r\n'
    + '                $WebRelUrl = $Op.WebUrl -replace [regex]::Escape($Op.SiteCollectionUrl), ""\r\n'
    + '                if (-not $WebRelUrl) { $WebRelUrl = "/" }\r\n'
    + '                $SPWeb = $SPSite.OpenWeb($WebRelUrl)\r\n'
    + '                $SPWeb.AllowUnsafeUpdates = $true\r\n'
    + '\r\n'
    + '                switch ($Op.Action) {\r\n'
    + '                    "RemoveDirectUserPermission" {\r\n'
    + '                        $User = $SPWeb.EnsureUser($Op.PrincipalLoginName)\r\n'
    + '                        if ($Op.ListId -and $Op.ItemId) {\r\n'
    + '                            $List = $SPWeb.Lists[[Guid]$Op.ListId]\r\n'
    + '                            $Item = $List.GetItemById([int]$Op.ItemId)\r\n'
    + '                            $Item.RoleAssignments.Remove($User); $Item.Update()\r\n'
    + '                        } elseif ($Op.ListId) {\r\n'
    + '                            $List = $SPWeb.Lists[[Guid]$Op.ListId]\r\n'
    + '                            $List.RoleAssignments.Remove($User); $List.Update()\r\n'
    + '                        } else {\r\n'
    + '                            $SPWeb.RoleAssignments.Remove($User); $SPWeb.Update()\r\n'
    + '                        }\r\n'
    + '                    }\r\n'
    + '                    "RemoveSharePointGroupAssignment" {\r\n'
    + '                        $Group = $SPWeb.SiteGroups[$Op.SharePointGroupName]\r\n'
    + '                        if ($Op.ListId -and $Op.ItemId) {\r\n'
    + '                            $List = $SPWeb.Lists[[Guid]$Op.ListId]\r\n'
    + '                            $Item = $List.GetItemById([int]$Op.ItemId)\r\n'
    + '                            $Item.RoleAssignments.Remove($Group); $Item.Update()\r\n'
    + '                        } elseif ($Op.ListId) {\r\n'
    + '                            $List = $SPWeb.Lists[[Guid]$Op.ListId]\r\n'
    + '                            $List.RoleAssignments.Remove($Group); $List.Update()\r\n'
    + '                        } else {\r\n'
    + '                            $SPWeb.RoleAssignments.Remove($Group); $SPWeb.Update()\r\n'
    + '                        }\r\n'
    + '                    }\r\n'
    + '                    "RemoveDomainGroupAssignment" {\r\n'
    + '                        $DGroup = $SPWeb.EnsureUser($Op.PrincipalLoginName)\r\n'
    + '                        if ($Op.ListId -and $Op.ItemId) {\r\n'
    + '                            $List = $SPWeb.Lists[[Guid]$Op.ListId]\r\n'
    + '                            $Item = $List.GetItemById([int]$Op.ItemId)\r\n'
    + '                            $Item.RoleAssignments.Remove($DGroup); $Item.Update()\r\n'
    + '                        } elseif ($Op.ListId) {\r\n'
    + '                            $List = $SPWeb.Lists[[Guid]$Op.ListId]\r\n'
    + '                            $List.RoleAssignments.Remove($DGroup); $List.Update()\r\n'
    + '                        } else {\r\n'
    + '                            $SPWeb.RoleAssignments.Remove($DGroup); $SPWeb.Update()\r\n'
    + '                        }\r\n'
    + '                    }\r\n'
    + '                    "RestoreInheritance" {\r\n'
    + '                        if ($Op.ListId -and $Op.ItemId) {\r\n'
    + '                            $List = $SPWeb.Lists[[Guid]$Op.ListId]\r\n'
    + '                            $Item = $List.GetItemById([int]$Op.ItemId)\r\n'
    + '                            $Item.ResetRoleInheritance(); $Item.Update()\r\n'
    + '                        } elseif ($Op.ListId) {\r\n'
    + '                            $List = $SPWeb.Lists[[Guid]$Op.ListId]\r\n'
    + '                            $List.ResetRoleInheritance(); $List.Update()\r\n'
    + '                        } else {\r\n'
    + '                            $SPWeb.ResetRoleInheritance(); $SPWeb.Update()\r\n'
    + '                        }\r\n'
    + '                    }\r\n'
    + '                }\r\n'
    + '\r\n'
    + '                Write-Host "  [OK] Wykonano: $($Op.Action)" -ForegroundColor Green\r\n'
    + '                $Success++\r\n'
    + '            } finally {\r\n'
    + '                if ($SPWeb) { $SPWeb.AllowUnsafeUpdates = $false; $SPWeb.Dispose() }\r\n'
    + '                if ($SPSite) { $SPSite.Dispose() }\r\n'
    + '            }\r\n'
    + '        }\r\n'
    + '    } catch {\r\n'
    + '        Write-Host "  [BLAD] $_" -ForegroundColor Red\r\n'
    + '        $Errors++\r\n'
    + '    }\r\n'
    + '}\r\n'
    + '\r\n'
    + 'Write-Host ""\r\n'
    + 'Write-Host "============ PODSUMOWANIE ============" -ForegroundColor Cyan\r\n'
    + 'Write-Host "  Sukces  : $Success" -ForegroundColor Green\r\n'
    + 'Write-Host "  Bledy   : $Errors" -ForegroundColor $(if ($Errors -gt 0) { "Red" } else { "Green" })\r\n'
    + 'Write-Host "  Pominieto: $Skipped" -ForegroundColor Yellow\r\n'
    + 'Stop-Transcript\r\n';
}

function exportRemediationPlan() {
  if (App.selectedRows.length === 0) {
    showToast('warning', 'Brak wybranych elementów', 'Zaznacz elementy przed eksportem.');
    return;
  }

  var plan = App.selectedRows.map(function(row) {
    return {
      ObjectId: row.obj.ObjectId,
      ObjectType: row.obj.ObjectType,
      FullUrl: row.obj.FullUrl || '',
      ServerRelativeUrl: row.obj.ServerRelativeUrl || '',
      SiteCollectionUrl: row.obj.SiteCollectionUrl || '',
      WebUrl: row.obj.WebUrl || '',
      ListId: row.obj.ListId || '',
      ItemId: row.obj.ItemId || '',
      PrincipalLoginName: row.assignment ? (row.assignment.LoginName || '') : '',
      PrincipalDisplayName: row.assignment ? (row.assignment.DisplayName || '') : '',
      PrincipalType: row.assignment ? (row.assignment.PrincipalType || '') : '',
      PermissionLevels: row.assignment ? (row.assignment.PermissionLevels || []) : [],
      SourceType: row.assignment ? (row.assignment.SourceType || '') : '',
      ExportedAt: new Date().toISOString()
    };
  });

  downloadText(JSON.stringify(plan, null, 2), 'RemediationPlan_' + new Date().toISOString().slice(0, 10) + '.json', 'application/json');
}

function normalizeSpUrl(url) {
  return (url || '').trim().replace(/\/$/, '');
}

function getUrlOrigin(url) {
  try {
    return new URL(url, window.location.href).origin;
  } catch (err) {
    return '';
  }
}

function ensureSameOriginSharePointTarget(url) {
  var normalizedUrl = normalizeSpUrl(url);
  var pageOrigin = window.location.origin || '';
  var targetOrigin = getUrlOrigin(normalizedUrl);

  if (!normalizedUrl || !targetOrigin) {
    throw new Error('Nieprawidlowy URL witryny SharePoint.');
  }
  if (!pageOrigin || pageOrigin === 'null') {
    throw new Error('Raport musi byc otwarty z witryny SharePoint, a nie z lokalnego pliku, aby uzyc REST API.');
  }
  if (pageOrigin !== targetOrigin) {
    throw new Error('Wybrana witryna SharePoint ma inny origin (' + targetOrigin + ') niz raport (' + pageOrigin + '). Otworz raport z tej samej aplikacji web albo uzyj skryptu PS1.');
  }
}

function getOperationContext(op, fallbackSiteUrl) {
  var siteUrl = normalizeSpUrl((op && op.SiteCollectionUrl) || fallbackSiteUrl || '');
  var webUrl = normalizeSpUrl((op && op.WebUrl) || (op && op.SiteCollectionUrl) || fallbackSiteUrl || '');

  if (op && op.ObjectType === 'WebApplication') {
    throw new Error('Bezposrednia remediacja REST nie obsluguje obiektow WebApplication. Uzyj skryptu PS1.');
  }
  if (!siteUrl && webUrl) siteUrl = webUrl;
  if (!webUrl) webUrl = siteUrl;
  if (!siteUrl || !webUrl) {
    throw new Error('Brak URL kontekstowego dla wybranego obiektu.');
  }

  ensureSameOriginSharePointTarget(siteUrl);
  ensureSameOriginSharePointTarget(webUrl);

  return {
    siteUrl: siteUrl,
    webUrl: webUrl
  };
}

function autoConnectDetectedSharePoint() {
  if (!SpContext.detected || !SpContext.siteUrl || SpContext.autoConnectAttempted) {
    return;
  }

  SpContext.autoConnectAttempted = true;
  testSpConnection({ siteUrl: SpContext.siteUrl, silent: true }).catch(function() {});
}

// ---- SharePoint REST API ----

function detectSharePointMode() {
  var siteUrlInput = document.getElementById('spSiteUrl');
  SpContext.connected = false;

  if (window._spPageContextInfo && window._spPageContextInfo.siteAbsoluteUrl) {
    SpContext.siteUrl = normalizeSpUrl(window._spPageContextInfo.siteAbsoluteUrl);
    SpContext.detected = true;
    if (siteUrlInput) siteUrlInput.value = SpContext.siteUrl;
    setSpStatus('detected', 'Auto-wykryto z _spPageContextInfo: ' + SpContext.siteUrl);
    return;
  }

  if (App.data && App.data.Objects) {
    var href = window.location.href.toLowerCase();
    var candidates = [];
    App.data.Objects.forEach(function(obj) {
      if (obj.ObjectType === 'SiteCollection' && obj.SiteCollectionUrl) {
        candidates.push(obj.SiteCollectionUrl);
      } else if (obj.ObjectType === 'WebApplication' && obj.WebApplicationUrl) {
        candidates.push(obj.WebApplicationUrl);
      }
    });
    for (var i = 0; i < candidates.length; i++) {
      if (href.indexOf(candidates[i].toLowerCase()) === 0) {
        SpContext.siteUrl = normalizeSpUrl(candidates[i]);
        SpContext.detected = true;
        if (siteUrlInput) siteUrlInput.value = SpContext.siteUrl;
        setSpStatus('detected', 'Auto-wykryto z URL strony: ' + SpContext.siteUrl);
        return;
      }
    }
  }

  SpContext.detected = false;
  SpContext.siteUrl = '';

  if (App.data && App.data.Objects) {
    var firstSC = App.data.Objects.find(function(o) { return o.ObjectType === 'SiteCollection'; });
    if (firstSC && firstSC.SiteCollectionUrl && siteUrlInput && !siteUrlInput.value) {
      siteUrlInput.placeholder = firstSC.SiteCollectionUrl;
    }
  }
  setSpStatus('unknown', 'Podaj URL witryny SharePoint i przetestuj polaczenie');
}

function setSpStatus(state, text) {
  var el = document.getElementById('spConnectionStatus');
  if (!el) return;
  el.className = 'sp-status sp-status-' + state;
  var icons = { unknown: '?', detected: '⟳', connected: '✓', error: '✕' };
  el.innerHTML = '<span class="sp-status-icon">' + (icons[state] || '?') + '</span>'
    + '<span class="sp-status-text">' + escapeHtml(text) + '</span>';
}

async function testSpConnection(options) {
  options = options || {};

  var input = document.getElementById('spSiteUrl');
  var url = normalizeSpUrl(options.siteUrl || (input ? input.value : ''));
  var btn = document.getElementById('btnTestSpConnection');
  var execBtn = document.getElementById('btnExecuteOnSP');

  if (!url) {
    if (!options.silent) {
      showToast('warning', 'Brak URL', 'Podaj URL witryny SharePoint.');
    }
    return false;
  }

  if (input) input.value = url;
  if (btn) btn.disabled = true;

  try {
    ensureSameOriginSharePointTarget(url);
    setSpStatus('detected', 'Testuje polaczenie...');
    await getSpFormDigest(url, true);
    SpContext.siteUrl = url;
    SpContext.connected = true;
    if (execBtn) execBtn.disabled = false;
    setSpStatus('connected', 'Polaczono: ' + url);
    if (!options.silent) {
      showToast('success', 'Polaczono', 'SharePoint REST API dostepne.');
    }
    return true;
  } catch (err) {
    SpContext.connected = false;
    if (execBtn) execBtn.disabled = true;
    setSpStatus('error', 'Blad polaczenia: ' + (err.message || String(err)));
    if (!options.silent) {
      showToast('error', 'Blad polaczenia', err.message || String(err));
    }
    throw err;
  } finally {
    if (btn) btn.disabled = false;
  }
}

async function getSpFormDigest(siteUrl, forceRefresh) {
  var normalizedSiteUrl = normalizeSpUrl(siteUrl);
  var cached = SpContext.digestCache[normalizedSiteUrl];

  if (!forceRefresh && cached && Date.now() < cached.expiry) {
    return cached.value;
  }

  if (!forceRefresh && window._spPageContextInfo && window._spPageContextInfo.formDigestValue) {
    var ctxUrl = normalizeSpUrl(window._spPageContextInfo.webAbsoluteUrl || window._spPageContextInfo.siteAbsoluteUrl || '');
    if (ctxUrl && ctxUrl === normalizedSiteUrl) {
      SpContext.digestCache[normalizedSiteUrl] = {
        value: window._spPageContextInfo.formDigestValue,
        expiry: Date.now() + 25 * 60 * 1000
      };
      return window._spPageContextInfo.formDigestValue;
    }
  }

  var resp = await fetch(normalizedSiteUrl + '/_api/contextinfo', {
    method: 'POST',
    credentials: 'include',
    headers: { 'Accept': 'application/json;odata=verbose', 'Content-Length': '0' }
  });
  if (!resp.ok) throw new Error('contextinfo HTTP ' + resp.status);
  var data = await resp.json();
  var info = data.d ? data.d.GetContextWebInformation : data.GetContextWebInformation;
  var ttl = (info.FormDigestTimeoutSeconds || 1500) - 30;

  SpContext.digestCache[normalizedSiteUrl] = {
    value: info.FormDigestValue,
    expiry: Date.now() + ttl * 1000
  };

  return info.FormDigestValue;
}

async function resolveSpPrincipalId(siteUrl, op) {
  var loginName = op ? (op.PrincipalLoginName || '') : '';
  var groupName = op ? (op.SharePointGroupName || op.PrincipalDisplayName || '') : '';
  var cacheKey = siteUrl + '|' + (loginName || ('spgroup:' + groupName));

  if (SpContext.principalCache[cacheKey] !== undefined) {
    return SpContext.principalCache[cacheKey];
  }

  if (loginName) {
    var resp = await fetch(siteUrl + "/_api/web/siteusers(@v)?@v='" + encodeURIComponent(loginName) + "'", {
      credentials: 'include',
      headers: { 'Accept': 'application/json;odata=verbose' }
    });

    if (resp.ok) {
      var data = await resp.json();
      var id = data.d ? data.d.Id : (data.Id || null);
      if (id === null || id === undefined) {
        throw new Error('Brak Id dla: ' + loginName);
      }
      SpContext.principalCache[cacheKey] = id;
      return id;
    }

    if (!groupName) {
      throw new Error('Nie znaleziono uzytkownika lub grupy: ' + loginName + ' (HTTP ' + resp.status + ')');
    }
  }

  if (groupName) {
    var groupResp = await fetch(siteUrl + "/_api/web/sitegroups/getByName('" + encodeURIComponent(groupName) + "')?$select=Id", {
      credentials: 'include',
      headers: { 'Accept': 'application/json;odata=verbose' }
    });

    if (!groupResp.ok) {
      throw new Error('Nie znaleziono grupy SharePoint: ' + groupName + ' (HTTP ' + groupResp.status + ')');
    }

    var groupData = await groupResp.json();
    var groupId = groupData.d ? groupData.d.Id : (groupData.Id || null);
    if (groupId === null || groupId === undefined) {
      throw new Error('Brak Id dla grupy SharePoint: ' + groupName);
    }

    SpContext.principalCache[cacheKey] = groupId;
    return groupId;
  }

  throw new Error('Brak danych principal do remediacji.');
}

async function executeSpOperation(context, op, digest) {
  var apiBase = context.webUrl + '/_api/web';
  var headersPost = {
    'Accept': 'application/json;odata=verbose',
    'X-RequestDigest': digest,
    'Content-Type': 'application/json;odata=verbose'
  };

  var action = op.Action;
  var listId = op.ListId ? op.ListId.replace(/^\{|\}$/g, '') : '';
  var itemId = op.ItemId;

  if (action === 'RestoreInheritance') {
    var endpoint;
    if (listId && itemId) {
      endpoint = apiBase + "/lists(guid'" + listId + "')/items(" + itemId + ")/resetroleinheritance()";
    } else if (listId) {
      endpoint = apiBase + "/lists(guid'" + listId + "')/resetroleinheritance()";
    } else {
      endpoint = apiBase + '/resetroleinheritance()';
    }
    var r = await fetch(endpoint, { method: 'POST', credentials: 'include', headers: headersPost });
    if (!r.ok) throw new Error('HTTP ' + r.status + ' dla resetroleinheritance');
    return;
  }

  if (action === 'RemoveDirectUserPermission' || action === 'RemoveSharePointGroupAssignment' || action === 'RemoveDomainGroupAssignment') {
    var principalId = await resolveSpPrincipalId(context.siteUrl, op);
    var delEndpoint;
    if (listId && itemId) {
      delEndpoint = apiBase + "/lists(guid'" + listId + "')/items(" + itemId + ")/roleAssignments/getByPrincipalId(" + principalId + ")/deleteObject()";
    } else if (listId) {
      delEndpoint = apiBase + "/lists(guid'" + listId + "')/roleAssignments/getByPrincipalId(" + principalId + ")/deleteObject()";
    } else {
      delEndpoint = apiBase + '/roleAssignments/getByPrincipalId(' + principalId + ')/deleteObject()';
    }
    var r2 = await fetch(delEndpoint, { method: 'POST', credentials: 'include', headers: headersPost });
    if (!r2.ok) throw new Error('HTTP ' + r2.status + ' dla deleteObject');
    return;
  }

  throw new Error('Nieznana akcja: ' + action);
}

function showSpProgress(total) {
  var panel = document.getElementById('spExecutionPanel');
  var log = document.getElementById('spResultsLog');
  var summary = document.getElementById('spSummary');
  if (panel) panel.style.display = '';
  if (log) log.innerHTML = '';
  if (summary) summary.innerHTML = '';
  updateSpProgress(0, total, '');
}

function updateSpProgress(current, total, label) {
  var fill = document.getElementById('spProgressFill');
  var text = document.getElementById('spProgressText');
  var pct = total > 0 ? Math.round(current / total * 100) : 0;
  if (fill) fill.style.width = pct + '%';
  if (text) text.textContent = current + ' / ' + total + (label ? ' — ' + label : '');
}

function logSpResult(op, status, error) {
  var log = document.getElementById('spResultsLog');
  if (!log) return;
  var cls = status === 'success' ? 'sp-log-ok' : (status === 'dryrun' ? 'sp-log-dry' : 'sp-log-err');
  var icon = status === 'success' ? '✓' : (status === 'dryrun' ? '○' : '✕');
  var msg = status === 'error' ? ' — ' + escapeHtml(error || '') : '';
  var line = document.createElement('div');
  line.className = 'sp-log-line ' + cls;
  line.textContent = icon + ' [' + (op.Action || '') + '] ' + (op.FullUrl || op.ServerRelativeUrl || op.ObjectId || '') + (op.PrincipalLoginName ? ' / ' + op.PrincipalLoginName : '') + msg;
  log.appendChild(line);
  log.scrollTop = log.scrollHeight;
}

function finishSpProgress(results, isDryRun) {
  var summary = document.getElementById('spSummary');
  if (!summary) return;
  var ok = results.filter(function(r) { return r.status === 'success' || r.status === 'dryrun'; }).length;
  var err = results.filter(function(r) { return r.status === 'error'; }).length;
  var mode = isDryRun ? 'DRY-RUN' : 'LIVE';
  var summaryHtml = '<strong>Podsumowanie [' + mode + ']:</strong> '+
    '<span class="sp-sum-ok">' + ok + ' operacji wykonanych</span>, '+
    (err > 0 ? '<span class="sp-sum-err">' + err + ' bledow</span>' : '<span class="sp-sum-ok">0 bledow</span>');
  if (!isDryRun && ok > 0) {
    summaryHtml += '<div class="sp-rescan-section">'+
      '<strong>Remediacja zakonczona.</strong> Odswiez dane raportu aby zobaczyc aktualne uprawnienia:'+
      ' <button class="btn btn-sm btn-primary" onclick="reloadReportData()" title="Pobierz nowy data.js i przelicz raport">Odswiez dane raportu</button>'+
      ' <button class="btn btn-sm" onclick="downloadRescanScript()" title="Pobierz skrypt PS1 do ponownego skanowania">Pobierz skrypt Rescan</button>'+
      '</div>';
  }
  summary.innerHTML = summaryHtml;
}

function reloadReportData() {
  showToast('info', 'Ladowanie danych...', 'Pobieranie nowego data.js...');
  var scriptTag = document.createElement('script');
  scriptTag.src = './data.js?v=' + Date.now();
  scriptTag.onload = function() {
    try {
      if (!window.SCAN_DATA) throw new Error('Brak danych SCAN_DATA w nowym data.js');
      App.data = window.SCAN_DATA;
      buildObjectIndexes();
      buildFlatRows();
      buildGroupedRows();
      buildUniqueListsForFilters();
      initDashboard();
      App.filteredGroupedRows = App.groupedRows.slice();
      App.filteredRows = App.flatRows.slice();
      rebuildTable();
      if (window.DIFF_DATA) { App.diffData = window.DIFF_DATA; initChangelogTab(); }
      showToast('success', 'Dane zaktualizowane', 'Raport zaladowal nowe dane ze skanowania.');
    } catch(err) {
      showToast('error', 'Blad odswiezania', err.message || String(err));
    }
  };
  scriptTag.onerror = function() {
    showToast('error', 'Blad odswiezania', 'Nie udalo sie pobrac data.js. Czy raport jest otwarty z serwera SharePoint?');
  };
  document.head.appendChild(scriptTag);
}

function downloadRescanScript() {
  var serverName = window.REPORT_SERVER || 'SERVER';
  var ps1 = '#Requires -Version 5.1\r\n'+
    '# Skrypt wygenerowany przez SharePoint Permission Analyzer\r\n'+
    '# Uruchom na serwerze SharePoint (' + serverName + ') jako Farm Administrator\r\n'+
    '# Ponownie uruchomi skanowanie po zakonczeniu remediacji\r\n\r\n'+
    '# Dostosuj sciezke do swoich ustawien:\r\n'+
    ' = "C:\\Path\\To\\Start-PermissionScan.ps1"\r\n\r\n'+
    '# Opcjonalnie: URL biblioteki SP do wgrania nowego raportu\r\n'+
    ' = ""\r\n\r\n'+
    'if (Test-Path ) {\r\n'+
    '    Write-Host "Uruchamianie skanowania: " -ForegroundColor Cyan\r\n'+
    '    if () {\r\n'+
    '        &  -SharePointLibraryUrl \r\n'+
    '    } else {\r\n'+
    '        & \r\n'+
    '    }\r\n'+
    '    Write-Host "Skanowanie zakonczone. Odswiez raport w przegladarce." -ForegroundColor Green\r\n'+
    '} else {\r\n'+
    '    Write-Error "Nie znaleziono skryptu skanowania: "\r\n'+
    '    Write-Host "Dostosuj zmienna ScanScriptPath i uruchom ponownie."\r\n'+
    '}\r\n';
  downloadText(ps1, 'Rescan-After-Remediation.ps1', 'text/plain;charset=utf-8;');
  showToast('info', 'Skrypt pobrany', 'Dostosuj sciezke ScanScriptPath i uruchom na serwerze SharePoint.');
}


async function executeOnSharePoint() {
  if (!App.selectedRows || App.selectedRows.length === 0) {
    showToast('warning', 'Brak wyboru', 'Zaznacz elementy w tabeli uprawnien.'); return;
  }
  var actionEl = document.getElementById('remediationAction');
  var action = actionEl ? actionEl.value : '';
  if (!action) { showToast('warning', 'Brak akcji', 'Wybierz akcje remediacyjna.'); return; }
  if (!SpContext.connected) { showToast('warning', 'Brak polaczenia', 'Najpierw przetestuj polaczenie z SharePoint.'); return; }

  var isDryRun = (document.getElementById('dryRunToggle') || {}).checked !== false;
  var siteUrl = normalizeSpUrl(SpContext.siteUrl);

  var plan = App.selectedRows.map(function(row) {
    return {
      Action: action,
      ObjectId: row.obj.ObjectId,
      ObjectType: row.obj.ObjectType,
      FullUrl: row.obj.FullUrl || '',
      ServerRelativeUrl: row.obj.ServerRelativeUrl || '',
      SiteCollectionUrl: row.obj.SiteCollectionUrl || '',
      WebUrl: row.obj.WebUrl || '',
      ListId: row.obj.ListId || '',
      ItemId: row.obj.ItemId || '',
      PrincipalLoginName: row.assignment ? (row.assignment.LoginName || '') : '',
      PrincipalDisplayName: row.assignment ? (row.assignment.DisplayName || '') : '',
      SharePointGroupName: row.assignment && row.assignment.PrincipalType === 'SharePointGroup'
        ? (row.assignment.DisplayName || row.assignment.LoginName || '')
        : ''
    };
  });

  showSpProgress(plan.length);
  var results = [];
  var execBtn = document.getElementById('btnExecuteOnSP');
  if (execBtn) execBtn.disabled = true;

  for (var i = 0; i < plan.length; i++) {
    var op = plan[i];
    var label = op.FullUrl || op.ServerRelativeUrl || op.ObjectId || '';
    updateSpProgress(i, plan.length, label);

    try {
      var opContext = getOperationContext(op, siteUrl);

      if (isDryRun) {
        var dr = { status: 'dryrun', op: op };
        results.push(dr);
        logSpResult(op, 'dryrun', null);
        continue;
      }

      var digest = await getSpFormDigest(opContext.webUrl, false);
      await executeSpOperation(opContext, op, digest);
      results.push({ status: 'success', op: op });
      logSpResult(op, 'success', null);
    } catch (err) {
      results.push({ status: 'error', op: op, error: err.message });
      logSpResult(op, 'error', err.message);
    }
  }

  updateSpProgress(plan.length, plan.length, 'Gotowe');
  finishSpProgress(results, isDryRun);
  if (execBtn) execBtn.disabled = false;
  showToast(results.some(function(r) { return r.status === 'error'; }) ? 'warning' : 'success',
    isDryRun ? 'DRY-RUN zakonczony' : 'Remediacja zakonczona',
    plan.length + ' operacji przetworzono');
}

function generateDeployScript() {
  var urlInput = document.getElementById('deployTargetUrl');
  var targetUrl = urlInput ? urlInput.value.trim() : '';
  if (!targetUrl) { showToast('warning', 'Brak URL', 'Podaj URL docelowej biblioteki SharePoint.'); return; }

  var reportFolder = escapeHtml(targetUrl);
  var ps1 = '#Requires -Version 3.0\r\n'
    + '<#\r\n'
    + '.SYNOPSIS\r\n'
    + '    Wdrazanie raportu uprawnien SharePoint do biblioteki dokumentow.\r\n'
    + '.DESCRIPTION\r\n'
    + '    Wgrywa wszystkie pliki z lokalnego folderu raportu do podanej biblioteki SP.\r\n'
    + '    Wymaga uprawnien do wybranej biblioteki.\r\n'
    + '#>\r\n'
    + '[CmdletBinding(SupportsShouldProcess)]\r\n'
    + 'param(\r\n'
    + '    [Parameter(Mandatory=$false)]\r\n'
    + '    [string]$TargetLibraryUrl = "' + targetUrl.replace(/"/g, '\\"') + '",\r\n'
    + '    [Parameter(Mandatory=$false)]\r\n'
    + '    [string]$ReportFolder = $PSScriptRoot,\r\n'
    + '    [switch]$DryRun = $false\r\n'
    + ')\r\n'
    + '\r\n'
    + 'Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue\r\n'
    + '\r\n'
    + 'function Upload-ReportToSP {\r\n'
    + '    param([string]$LibraryUrl, [string]$LocalFolder)\r\n'
    + '\r\n'
    + '    # Wyodrebnij URL witryny i relatywna sciezke biblioteki\r\n'
    + '    $Uri = [Uri]$LibraryUrl\r\n'
    + '    $SiteUrl = $Uri.GetLeftPart([UriPartial]::Authority) + ($Uri.AbsolutePath -replace "/[^/]+$", "")\r\n'
    + '    $LibRelUrl = $Uri.AbsolutePath\r\n'
    + '\r\n'
    + '    $SPSite = $null; $SPWeb = $null\r\n'
    + '    try {\r\n'
    + '        $SPSite = New-Object Microsoft.SharePoint.SPSite($SiteUrl)\r\n'
    + '        $SPWeb = $SPSite.OpenWeb("/")\r\n'
    + '        $SPWeb.AllowUnsafeUpdates = $true\r\n'
    + '        $Lib = $SPWeb.GetFolder($LibRelUrl)\r\n'
    + '        if (-not $Lib.Exists) { throw "Biblioteka nie istnieje: $LibRelUrl" }\r\n'
    + '\r\n'
    + '        # Wgraj wszystkie pliki z folderu (rekurencyjnie)\r\n'
    + '        $Files = Get-ChildItem -Path $LocalFolder -File -Recurse\r\n'
    + '        $Total = $Files.Count; $Idx = 0\r\n'
    + '        foreach ($File in $Files) {\r\n'
    + '            $Idx++\r\n'
    + '            $RelPath = $File.FullName.Substring($LocalFolder.Length).TrimStart([char[]]"\\/")\r\n'
    + '            $RelPath = $RelPath.Replace("\\", "/")\r\n'
    + '            $DestRelUrl = $LibRelUrl.TrimEnd("/") + "/" + $RelPath\r\n'
    + '            Write-Host "[$Idx/$Total] $RelPath" -NoNewline\r\n'
    + '            if ($DryRun) { Write-Host " [DRY-RUN]" -ForegroundColor Yellow; continue }\r\n'
    + '            $Bytes = [System.IO.File]::ReadAllBytes($File.FullName)\r\n'
    + '            $null = $Lib.Files.Add($DestRelUrl, $Bytes, $true)\r\n'
    + '            Write-Host " OK" -ForegroundColor Green\r\n'
    + '        }\r\n'
    + '        Write-Host "Wdrozono $Total plik(ow) do $LibRelUrl" -ForegroundColor Cyan\r\n'
    + '    } finally {\r\n'
    + '        if ($SPWeb) { $SPWeb.Dispose() }\r\n'
    + '        if ($SPSite) { $SPSite.Dispose() }\r\n'
    + '    }\r\n'
    + '}\r\n'
    + '\r\n'
    + 'Upload-ReportToSP -LibraryUrl $TargetLibraryUrl -LocalFolder $ReportFolder\r\n';

  downloadText(ps1, 'Deploy-Report.ps1', 'text/plain');
}

/* ============================================================
   SEKCJA 11: NAWIGACJA
   ============================================================ */

function initNavigation() {
  document.querySelectorAll('.nav-tab').forEach(function(tab) {
    tab.addEventListener('click', function() {
      var tabName = this.getAttribute('data-tab');
      switchTab(tabName);
    });
  });
}

function switchTab(tabName) {
  // Ukryj wszystkie taby
  document.querySelectorAll('.tab-content').forEach(function(el) {
    el.style.display = 'none';
    el.classList.remove('active');
  });

  // Odznacz wszystkie nav-tab
  document.querySelectorAll('.nav-tab').forEach(function(el) {
    el.classList.remove('active');
  });

  // Pokaż wybrany tab
  var tabContent = document.getElementById('tab-' + tabName);
  if (tabContent) {
    tabContent.style.display = '';
    tabContent.classList.add('active');
  }

  // Zaznacz aktywny nav-tab
  var navTab = document.querySelector('.nav-tab[data-tab="' + tabName + '"]');
  if (navTab) navTab.classList.add('active');

  // Odśwież wykresy jeśli dashboard
  if (tabName === 'dashboard' && Object.keys(App.charts).length > 0) {
    Object.values(App.charts).forEach(function(chart) {
      if (chart && chart.update) chart.update();
    });
  }

  // Odśwież drzewo jeśli tree tab
  if (tabName === 'tree' && App.treeInstance) {
    setTimeout(function() {
      if (App.treeInstance) App.treeInstance.refresh();
    }, 100);
  }

  // Odśwież panel remediacji
  if (tabName === 'remediation') {
    updateRemediationPanel();
  }
}

/* ============================================================
   SEKCJA 12: MOTYW JASNY/CIEMNY
   ============================================================ */

function initTheme() {
  var savedTheme = localStorage.getItem('sp-analyzer-theme') || 'light';
  applyTheme(savedTheme);
}

function initThemeToggle() {
  var btn = document.getElementById('themeToggle');
  if (btn) {
    btn.addEventListener('click', function() {
      var current = document.documentElement.getAttribute('data-theme');
      applyTheme(current === 'dark' ? 'light' : 'dark');
    });
  }
}

function applyTheme(theme) {
  document.documentElement.setAttribute('data-theme', theme);
  App.currentTheme = theme;
  localStorage.setItem('sp-analyzer-theme', theme);

  var iconSun = document.getElementById('iconSun');
  var iconMoon = document.getElementById('iconMoon');
  if (iconSun) iconSun.style.display = theme === 'dark' ? '' : 'none';
  if (iconMoon) iconMoon.style.display = theme === 'light' ? '' : 'none';

  // Odśwież wykresy dla nowego motywu
  if (Object.keys(App.charts).length > 0) {
    setTimeout(function() {
      Object.values(App.charts).forEach(function(c) { if (c) c.destroy(); });
      App.charts = {};
      renderCharts();
    }, 50);
  }
}

/* ============================================================
   SEKCJA 13: NAGŁÓWEK RAPORTU
   ============================================================ */

function updateReportHeader() {
  var meta = App.data.ScanMetadata || {};
  var subtitle = document.getElementById('reportSubtitle');
  var scanInfo = document.getElementById('scanInfo');

  var farmName = meta.FarmName || 'SharePoint Farm';
  var scanDate = meta.ScanStartTime ? new Date(meta.ScanStartTime).toLocaleDateString('pl-PL') : '-';
  var scanTime = meta.ScanStartTime ? new Date(meta.ScanStartTime).toLocaleTimeString('pl-PL') : '-';

  if (subtitle) setTextSafe(subtitle, farmName + ' | Skan: ' + scanDate + ' ' + scanTime);
  if (scanInfo) setTextSafe(scanInfo, 'Serwer: ' + (meta.ScanServer || '-') + ' | Przez: ' + (meta.ScanUser || '-'));

  document.title = 'SP Permission Analyzer - ' + farmName;

  var titleEl = document.querySelector('.brand-title');
  // Nie zmieniaj tytułu - zostaw domyślny
}

/* ============================================================
   SEKCJA 14: EKSPORT
   ============================================================ */

function exportTableCsv() {
  var rows = App.filteredRows;
  if (rows.length === 0) {
    showToast('warning', 'Brak danych', 'Tabela jest pusta.');
    return;
  }

  var headers = ['ObjectId','ObjectType','WebApplicationUrl','SiteCollectionUrl','WebUrl','FullUrl','ServerRelativeUrl',
    'Title','Name','ListTitle','ListId','ItemId','FileLeafRef','IsHidden','IsSystem','HasUniquePermissions',
    'InheritsFromUrl','PrincipalType','LoginName','DisplayName','Email','SourceType','SourceName',
    'PermissionLevels','IsLimitedAccessOnly','IsSiteAdmin','IsOrphaned','IsUnresolved'];

  var csvRows = [headers.join(';')];

  rows.forEach(function(row) {
    var obj = row.obj;
    var a = row.assignment;
    var cells = [
      obj.ObjectId, obj.ObjectType, obj.WebApplicationUrl, obj.SiteCollectionUrl, obj.WebUrl,
      obj.FullUrl, obj.ServerRelativeUrl, obj.Title, obj.Name, obj.ListTitle, obj.ListId,
      obj.ItemId, obj.FileLeafRef, obj.IsHidden, obj.IsSystem, obj.HasUniquePermissions,
      obj.InheritsFromUrl,
      a ? a.PrincipalType : '', a ? a.LoginName : '', a ? a.DisplayName : '', a ? a.Email : '',
      a ? a.SourceType : '', a ? a.SourceName : '',
      a ? (a.PermissionLevels || []).join('|') : '',
      a ? a.IsLimitedAccessOnly : '', a ? a.IsSiteAdmin : '',
      a ? a.IsOrphaned : '', a ? a.IsUnresolved : ''
    ];
    csvRows.push(cells.map(function(c) { return '"' + String(c === null || c === undefined ? '' : c).replace(/"/g, '""') + '"'; }).join(';'));
  });

  downloadText(csvRows.join('\r\n'), 'PermissionReport_' + new Date().toISOString().slice(0, 10) + '.csv', 'text/csv;charset=utf-8;');
  showToast('success', 'Eksport CSV', 'Pobieranie ' + rows.length + ' wierszy...');
}

function exportTableJson() {
  var rows = App.filteredRows;
  var data = rows.map(function(row) {
    return { object: row.obj, assignment: row.assignment };
  });
  downloadText(JSON.stringify(data, null, 2), 'PermissionReport_' + new Date().toISOString().slice(0, 10) + '.json', 'application/json');
  showToast('success', 'Eksport JSON', 'Pobieranie ' + rows.length + ' wierszy...');
}

function exportDashboardJson() {
  var data = {
    scanMetadata: App.data.ScanMetadata,
    statistics: App.data.Statistics
  };
  downloadText(JSON.stringify(data, null, 2), 'Dashboard_' + new Date().toISOString().slice(0, 10) + '.json', 'application/json');
}

function downloadText(content, filename, mimeType) {
  var BOM = mimeType.indexOf('csv') !== -1 ? '\uFEFF' : '';
  var blob = new Blob([BOM + content], { type: mimeType });
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(function() {
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, 100);
}

/* ============================================================
   SEKCJA 15: NARZĘDZIA - IKONY I BADGE
   ============================================================ */

function getObjectTypeIcon(objectType) {
  var icons = {
    'WebApplication': '<svg class="obj-type-icon obj-icon-webapp" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:14px;height:14px;display:inline"><rect x="2" y="3" width="20" height="14" rx="2"/><line x1="8" y1="21" x2="16" y2="21"/><line x1="12" y1="17" x2="12" y2="21"/></svg>',
    'SiteCollection': '<svg class="obj-type-icon obj-icon-site" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:14px;height:14px;display:inline"><path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/></svg>',
    'Web':            '<svg class="obj-type-icon obj-icon-web" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:14px;height:14px;display:inline"><circle cx="12" cy="12" r="10"/><line x1="2" y1="12" x2="22" y2="12"/><path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"/></svg>',
    'List':           '<svg class="obj-type-icon obj-icon-list" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:14px;height:14px;display:inline"><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/></svg>',
    'Library':        '<svg class="obj-type-icon obj-icon-library" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:14px;height:14px;display:inline"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></svg>',
    'Folder':         '<svg class="obj-type-icon obj-icon-folder" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:14px;height:14px;display:inline"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></svg>',
    'File':           '<svg class="obj-type-icon obj-icon-file" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:14px;height:14px;display:inline"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>',
    'ListItem':       '<svg class="obj-type-icon obj-icon-item" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:14px;height:14px;display:inline"><rect x="3" y="3" width="18" height="18" rx="2" ry="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="3" y1="15" x2="21" y2="15"/><line x1="9" y1="3" x2="9" y2="21"/></svg>'
  };
  return icons[objectType] || '<svg class="obj-type-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:14px;height:14px;display:inline"><circle cx="12" cy="12" r="2"/></svg>';
}

function getObjectTypeClass(objectType) {
  var map = {
    'WebApplication': 'webapp', 'SiteCollection': 'site', 'Web': 'web',
    'List': 'list', 'Library': 'library', 'Folder': 'folder',
    'File': 'file', 'ListItem': 'item'
  };
  return map[objectType] || 'web';
}

function getPrincipalTypeBadge(principalType) {
  var map = {
    'User':             '<span class="badge badge-direct">Użytkownik</span>',
    'SharePointGroup':  '<span class="badge badge-via-group">Grupa SP</span>',
    'DomainGroup':      '<span class="badge badge-via-ad">Grupa AD</span>',
    'SpecialPrincipal': '<span class="badge badge-limited">Specjalny</span>',
    'Claim':            '<span class="badge badge-limited">Claim</span>'
  };
  return map[principalType] || '<span class="badge">' + escapeHtml(principalType || '-') + '</span>';
}

function getPrincipalTypeIconStr(principalType) {
  var icons = {
    'User': '👤', 'SharePointGroup': '🏷', 'DomainGroup': '🏢',
    'SpecialPrincipal': '⭐', 'Claim': '🔑'
  };
  return icons[principalType] || '?';
}

function getSourceBadge(sourceType) {
  var map = {
    'Direct':             '<span class="badge badge-direct">Bezpośrednie</span>',
    'ViaSharePointGroup': '<span class="badge badge-via-group">Przez grupę SP</span>',
    'ViaDomainGroup':     '<span class="badge badge-via-ad">Przez grupę AD</span>',
    'Inherited':          '<span class="badge badge-inherited">Dziedziczone</span>'
  };
  return map[sourceType] || '<span class="badge">' + escapeHtml(sourceType || '-') + '</span>';
}

/* ============================================================
   SEKCJA 16: TOASTY (powiadomienia)
   ============================================================ */

function showToast(type, title, message, duration) {
  duration = duration || 4000;
  var container = document.getElementById('toastContainer');
  if (!container) return;

  var icons = { success: '✓', error: '✕', warning: '⚠', info: 'ℹ' };

  var toast = document.createElement('div');
  toast.className = 'toast toast-' + type;
  toast.innerHTML = '<span class="toast-icon">' + (icons[type] || 'ℹ') + '</span>'
    + '<div class="toast-body">'
    + '<div class="toast-title">' + escapeHtml(title) + '</div>'
    + (message ? '<div class="toast-message">' + escapeHtml(message) + '</div>' : '')
    + '</div>';

  container.appendChild(toast);

  setTimeout(function() {
    toast.style.opacity = '0';
    toast.style.transition = 'opacity 0.3s ease';
    setTimeout(function() {
      if (toast.parentNode) toast.parentNode.removeChild(toast);
    }, 300);
  }, duration);

  toast.addEventListener('click', function() {
    if (toast.parentNode) toast.parentNode.removeChild(toast);
  });
}

/* ============================================================
   SEKCJA 17: WERSJONOWANIE (CHANGELOG)
   ============================================================ */

function initChangelogTab() {
  var diff = window.DIFF_DATA;
  var history = window.SCAN_HISTORY;
  var container = document.getElementById('changelogContent');
  var meta = document.getElementById('changelogMeta');
  if (!container) return;

  var html = '';

  // Sekcja: historia skanowania (pokazuj zawsze jesli dostepna)
  if (history && history.length > 0) {
    html += renderScanHistory(history);
  }

  if (!diff || diff.IsFirstReport) {
    if (meta) meta.textContent = 'Brak poprzedniego raportu';
    if (!html) {
      html +=
        '<div style="padding:40px;text-align:center;color:var(--text-muted)">' +
        '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" style="width:48px;height:48px;margin:0 auto 16px;display:block">' +
        '<circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>' +
        '<p style="font-weight:600;margin-bottom:8px">Brak danych poprzedniego raportu</p>' +
        '<p>To jest pierwszy raport bazowy. Przy kolejnym skanowaniu ta zakladka pokaze<br>roznice w uprawnieniach wzgledem niniejszego raportu.</p>' +
        '</div>';
    }
    container.innerHTML = html;
    if (history && history.length > 0) { renderHistoryChart(history); }
    return;
  }

  var added   = (diff.AddedObjects   || []).length;
  var removed = (diff.RemovedObjects  || []).length;
  var changed = (diff.ChangedObjects  || []).length;

  if (meta) meta.textContent = 'Poprzedni: ' + (diff.PreviousGenerated || '-') + '   Aktualny: ' + (diff.CurrentGenerated || '-');

  // Separator miedzy historia a sekcja diff
  if (history && history.length > 0) {
    html += '<div style="padding:0 24px"><h3 style="margin:0;padding:20px 0 12px;border-top:1px solid var(--border)">Zestawienie zmian</h3></div>';
  }

  // Pasek porownania dat
  html += '<div style="padding:14px 24px;background:var(--bg-card);border-bottom:1px solid var(--border);display:flex;gap:32px;flex-wrap:wrap;align-items:center;margin-bottom:0">';
  html += '<div><span style="font-size:11px;color:var(--text-muted);text-transform:uppercase;letter-spacing:.04em">Poprzedni raport</span><br><strong style="font-size:14px">' + escapeHtml(diff.PreviousGenerated || '-') + '</strong></div>';
  html += '<div style="font-size:20px;color:var(--text-muted)">&#x2192;</div>';
  html += '<div><span style="font-size:11px;color:var(--text-muted);text-transform:uppercase;letter-spacing:.04em">Aktualny raport</span><br><strong style="font-size:14px">' + escapeHtml(diff.CurrentGenerated || '-') + '</strong></div>';
  html += '</div>';

  // Karty podsumowania
  html += '<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:16px;padding:20px 24px">';
  html += changelogCard('+' + added,   'Nowe obiekty',    '#107c10', added > 0   ? '#f0fff0' : '');
  html += changelogCard('-' + removed, 'Usuniete obiekty','#a4262c', removed > 0 ? '#fff0f0' : '');
  html += changelogCard('~' + changed, 'Zmiany uprawnien','#ca5010', changed > 0 ? '#fff8f0' : '');
  html += '</div>';

  if (added === 0 && removed === 0 && changed === 0) {
    html += '<div style="padding:32px;text-align:center;color:var(--text-muted)"><p>&#x2713; Brak zmian wzgledem poprzedniego raportu.</p></div>';
    container.innerHTML = html;
    if (history && history.length > 0) { renderHistoryChart(history); }
    return;
  }

  // Sekcja: zmiany uprawnien
  if (changed > 0) {
    html += '<div style="padding:0 24px 24px">';
    html += '<h3 style="margin:0 0 12px;padding-top:8px">Zmiany uprawnien (' + changed + ')</h3>';
    (diff.ChangedObjects || []).forEach(function(obj) {
      html += changelogObjectCard(obj);
    });
    html += '</div>';
  }

  // Sekcja: nowe obiekty
  if (added > 0) {
    html += changelogObjectsTable(diff.AddedObjects || [], 'Nowe obiekty (' + added + ')', '#107c10');
  }

  // Sekcja: usuniete obiekty
  if (removed > 0) {
    html += changelogObjectsTable(diff.RemovedObjects || [], 'Usuniete obiekty (' + removed + ')', '#a4262c');
  }

  container.innerHTML = html;
  if (history && history.length > 0) { renderHistoryChart(history); }
}


/* ---- Historia skanowania ---- */

function formatBytes(bytes) {
  if (!bytes || bytes === 0) return '0 B';
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / 1048576).toFixed(2) + ' MB';
}

function historyDelta(delta, isBytes) {
  if (delta === null || delta === undefined || delta === 0) return '';
  var positive = delta > 0;
  var color = positive ? '#c7302e' : '#107c10';
  if (isBytes) color = positive ? '#ca5010' : '#107c10';
  var abs = isBytes ? formatBytes(Math.abs(delta)) : Math.abs(delta);
  return ' <span style="font-size:11px;color:' + color + '">(' + (positive ? '+' : '-') + abs + ')</span>';
}

function renderScanHistory(history) {
  var html = '<div style="padding:0 24px 24px">';
  html += '<h3 style="margin:0 0 16px;padding-top:24px">Historia skanowania (ostatnie ' + history.length + ' wersji)</h3>';

  // Wykres trendu
  html += '<div style="background:var(--bg-card);border:1px solid var(--border);border-radius:8px;padding:16px 16px 8px;margin-bottom:20px">';
  html += '<div style="font-size:12px;font-weight:600;color:var(--text-muted);margin-bottom:8px;text-transform:uppercase;letter-spacing:.05em">Trend: obiekty / przypisania / rozmiar raportu</div>';
  html += '<canvas id="historyChart" height="110"></canvas>';
  html += '</div>';

  // Tabela porownan
  html += '<table class="detail-table">';
  html += '<thead><tr>';
  html += '<th>Data skanowania</th>';
  html += '<th style="text-align:right">Obiekty</th>';
  html += '<th style="text-align:right">Przypisania</th>';
  html += '<th style="text-align:right">Unikalne ACL</th>';
  html += '<th style="text-align:right">Web App</th>';
  html += '<th style="text-align:right">Site Collections</th>';
  html += '<th style="text-align:right">Rozmiar raportu</th>';
  html += '</tr></thead><tbody>';

  history.forEach(function(h, i) {
    var prev = i > 0 ? history[i - 1] : null;
    var isLast = i === history.length - 1;
    var rowStyle = isLast ? ' style="background:var(--bg-hover);font-weight:600"' : '';
    html += '<tr' + rowStyle + '>';

    var label = h.Generated || h.FolderName || '';
    html += '<td>' + escapeHtml(label) +
      (isLast ? ' <span style="font-size:10px;background:#0078d4;color:#fff;padding:1px 6px;border-radius:3px;margin-left:4px">najnowszy</span>' : '') +
      '</td>';

    var dObj  = prev ? (h.TotalObjectsScanned - prev.TotalObjectsScanned) : null;
    var dAsg  = prev ? (h.TotalAssignments - prev.TotalAssignments) : null;
    var dSize = prev ? (h.FolderSizeBytes - prev.FolderSizeBytes) : null;

    html += '<td style="text-align:right">' + (h.TotalObjectsScanned || 0) + historyDelta(dObj, false) + '</td>';
    html += '<td style="text-align:right">' + (h.TotalAssignments || 0) + historyDelta(dAsg, false) + '</td>';
    html += '<td style="text-align:right">' + (h.UniquePermissionsCount || 0) + '</td>';
    html += '<td style="text-align:right">' + (h.WebApplicationCount || 0) + '</td>';
    html += '<td style="text-align:right">' + (h.SiteCollectionCount || 0) + '</td>';
    html += '<td style="text-align:right">' + formatBytes(h.FolderSizeBytes || 0) + historyDelta(dSize, true) + '</td>';
    html += '</tr>';
  });

  html += '</tbody></table>';
  html += '</div>';
  return html;
}

function renderHistoryChart(history) {
  var ctx = document.getElementById('historyChart');
  if (!ctx || typeof Chart === 'undefined') return;
  if (App.charts.historyChart) { App.charts.historyChart.destroy(); delete App.charts.historyChart; }

  var isDark = document.documentElement.getAttribute('data-theme') === 'dark';
  var textColor = isDark ? '#cccccc' : '#323130';
  var gridColor = isDark ? 'rgba(255,255,255,.1)' : 'rgba(0,0,0,.07)';

  var labels = history.map(function(h) {
    var g = h.Generated || h.FolderName || '';
    return g.length >= 16 ? g.substring(0, 16) : g;
  });
  var sizeKb   = history.map(function(h) { return Math.round((h.FolderSizeBytes || 0) / 1024); });
  var objData  = history.map(function(h) { return h.TotalObjectsScanned || 0; });
  var asgData  = history.map(function(h) { return h.TotalAssignments || 0; });

  App.charts.historyChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [
        {
          label: 'Rozmiar raportu (KB)',
          data: sizeKb,
          backgroundColor: isDark ? 'rgba(96,180,240,.65)' : 'rgba(0,120,212,.65)',
          borderColor:      isDark ? '#60b4f0' : '#0078d4',
          borderWidth: 1,
          yAxisID: 'ySize'
        },
        {
          label: 'Obiekty',
          data: objData,
          type: 'line',
          borderColor: isDark ? '#81c784' : '#2e7d32',
          backgroundColor: 'transparent',
          pointRadius: 4,
          borderWidth: 2,
          tension: 0.2,
          yAxisID: 'yCount'
        },
        {
          label: 'Przypisania',
          data: asgData,
          type: 'line',
          borderColor: isDark ? '#ffb74d' : '#e65100',
          backgroundColor: 'transparent',
          pointRadius: 4,
          borderWidth: 2,
          tension: 0.2,
          yAxisID: 'yCount'
        }
      ]
    },
    options: {
      responsive: true,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { position: 'top', labels: { color: textColor, boxWidth: 14, padding: 10 } }
      },
      scales: {
        ySize: {
          type: 'linear', position: 'left',
          title: { display: true, text: 'KB', color: isDark ? '#60b4f0' : '#0078d4', font: { size: 11 } },
          ticks: { color: textColor },
          grid: { color: gridColor }
        },
        yCount: {
          type: 'linear', position: 'right',
          title: { display: true, text: 'Liczba', color: isDark ? '#ffb74d' : '#e65100', font: { size: 11 } },
          ticks: { color: textColor },
          grid: { drawOnChartArea: false }
        },
        x: {
          ticks: { color: textColor, maxRotation: 40, font: { size: 10 } },
          grid: { color: gridColor }
        }
      }
    }
  });
}
function changelogCard(value, label, color, bg) {
  return '<div style="background:' + (bg || 'var(--bg-card)') + ';border:1px solid var(--border);border-left:4px solid ' + color + ';border-radius:8px;padding:16px;text-align:center">' +
    '<div style="font-size:2rem;font-weight:700;color:' + color + ';line-height:1">' + escapeHtml(String(value)) + '</div>' +
    '<div style="color:var(--text-muted);font-size:12px;margin-top:4px">' + escapeHtml(label) + '</div>' +
    '</div>';
}

function changelogObjectCard(obj) {
  var html = '<div style="background:var(--bg-card);border:1px solid var(--border);border-radius:8px;overflow:hidden;margin-bottom:12px">';

  // Nagłówek obiektu
  html += '<div style="padding:10px 16px;background:var(--bg-hover);border-bottom:1px solid var(--border);display:flex;align-items:center;gap:8px;flex-wrap:wrap">';
  html += getObjectTypeIcon(obj.ObjectType);
  html += '<strong>' + escapeHtml(obj.Title || obj.ObjectType || '') + '</strong>';
  if (obj.ServerRelativeUrl) {
    html += '<span style="color:var(--text-muted);font-size:11px">' + escapeHtml(obj.ServerRelativeUrl) + '</span>';
  }
  if (obj.UniquePermissionsChanged) {
    var from = obj.OldHasUnique ? 'Unikatowe' : 'Dziedziczone';
    var to   = obj.NewHasUnique ? 'Unikatowe' : 'Dziedziczone';
    html += '<span style="margin-left:auto;font-size:11px;padding:2px 8px;border-radius:4px;background:var(--bg-card);border:1px solid var(--border)">' +
      escapeHtml(from) + ' → ' + escapeHtml(to) + '</span>';
  }
  html += '</div>';

  // Ciało karty z tabelami zmian
  html += '<div style="padding:12px 16px">';

  var addedA   = obj.AddedAssignments   || [];
  var removedA = obj.RemovedAssignments || [];

  if (addedA.length > 0) {
    html += '<div style="margin-bottom:10px">';
    html += '<div style="font-size:11px;font-weight:700;color:#107c10;margin-bottom:4px;text-transform:uppercase;letter-spacing:.04em">+ Dodane uprawnienia</div>';
    html += changelogAssignmentsTable(addedA, '#f0fff0');
    html += '</div>';
  }
  if (removedA.length > 0) {
    html += '<div>';
    html += '<div style="font-size:11px;font-weight:700;color:#a4262c;margin-bottom:4px;text-transform:uppercase;letter-spacing:.04em">− Usunięte uprawnienia</div>';
    html += changelogAssignmentsTable(removedA, '#fff0f0');
    html += '</div>';
  }

  html += '</div></div>';
  return html;
}

function changelogAssignmentsTable(assignments, rowBg) {
  var html = '<table class="detail-table"><thead><tr><th>Typ</th><th>Użytkownik / Grupa</th><th>Poziomy uprawnień</th><th>Źródło</th></tr></thead><tbody>';
  assignments.forEach(function(a) {
    html += '<tr style="background:' + rowBg + '">';
    html += '<td>' + escapeHtml(a.PrincipalType || '') + '</td>';
    html += '<td>' + escapeHtml(a.DisplayName || a.LoginName || '') + '</td>';
    html += '<td>' + escapeHtml((a.PermissionLevels || []).join(', ')) + '</td>';
    html += '<td>' + escapeHtml(a.SourceType === 'Group' ? (a.SourceName || 'Group') : (a.SourceType || '')) + '</td>';
    html += '</tr>';
  });
  html += '</tbody></table>';
  return html;
}

function changelogObjectsTable(objects, title, color) {
  var html = '<div style="padding:0 24px 24px">';
  html += '<h3 style="margin:0 0 12px;padding-top:8px;color:' + color + '">' + escapeHtml(title) + '</h3>';
  html += '<table class="detail-table"><thead><tr><th>Typ</th><th>Nazwa</th><th>URL</th><th>ACL</th></tr></thead><tbody>';
  objects.forEach(function(obj) {
    html += '<tr>';
    html += '<td>' + getObjectTypeIcon(obj.ObjectType) + ' ' + escapeHtml(obj.ObjectType || '') + '</td>';
    html += '<td>' + escapeHtml(obj.Title || obj.ObjectType || '') + '</td>';
    html += '<td style="font-size:11px;color:var(--text-muted)">' + escapeHtml(obj.ServerRelativeUrl || '') + '</td>';
    html += '<td>' + (obj.HasUniquePermissions
      ? '<span class="badge badge-unique">Unikatowe</span>'
      : '<span class="badge badge-inherited">Dziedziczone</span>') + '</td>';
    html += '</tr>';
  });
  html += '</tbody></table></div>';
  return html;
}

/* ============================================================
   SEKCJA 18: FALLBACK (gdy DataTables niedostępne)
   ============================================================ */

function initTableFallback() {
  // Prosta tabela bez DataTables
  var tbody = document.getElementById('tableBody');
  if (!tbody) return;
  var html = '';
  App.filteredGroupedRows.slice(0, 1000).forEach(function(gRow) { html += buildGroupedTableRow(gRow); });
  tbody.innerHTML = html;
  if (App.filteredGroupedRows.length > 1000) {
    showToast('info', 'Ograniczono widok', 'DataTables niedostępne. Wyświetlono pierwsze 1000 z ' + App.filteredRows.length + ' wierszy.');
  }
}





