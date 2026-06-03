// Logs Manager - Frontend
let currentFilters = {
  level: 'INFO',
  category: 'ALL',
  search: ''
};

let autoScroll = true;
let isPaused = false;
let stats = { DEBUG: 0, INFO: 0, SUCCESS: 0, WARN: 0, ERROR: 0 };
let totalLogs = 0;
let liveQueue = [];
let liveFlushTimer = null;

function invokeBridge(channel, payload) {
  if (window.electronAPI?.invoke) {
    return window.electronAPI.invoke(channel, payload);
  }
  if (window.api?.invoke) {
    return window.api.invoke(channel, payload);
  }
  throw new Error('Bridge IPC indisponible');
}

// Initialisation
document.addEventListener('DOMContentLoaded', () => {
  console.log('📱 Initialisation de la page logs...');
  loadHistory();
  setupListeners();
  setupFilterListeners();
});

async function loadHistory() {
  try {
    console.log('📥 Chargement de l\'historique des logs...');
    const history = await invokeBridge('api-get-log-history', currentFilters);
    const container = document.getElementById('logs-container');
    container.innerHTML = '';
    
    if (!history || history.length === 0) {
      container.innerHTML = `
        <div class="text-center text-muted py-5">
          <i class="bi bi-inbox" style="font-size: 48px;"></i>
          <p class="mt-3">Aucun log disponible</p>
        </div>
      `;
      updateStats();
      return;
    }
    
    // Réinitialiser les stats
    stats = { DEBUG: 0, INFO: 0, SUCCESS: 0, WARN: 0, ERROR: 0 };

    const fragment = document.createDocumentFragment();
    history.forEach(log => {
      const entry = createLogEntryElement(log);
      fragment.appendChild(entry);
      stats[log.level] = (stats[log.level] || 0) + 1;
    });
    container.appendChild(fragment);
    
    totalLogs = history.length;
    updateStats();
    scrollToBottom();
    console.log(`✅ ${history.length} logs chargés`);
  } catch (err) {
    console.error('❌ Erreur chargement historique:', err);
    const container = document.getElementById('logs-container');
    container.innerHTML = `
      <div class="text-center text-danger py-5">
        <i class="bi bi-exclamation-triangle" style="font-size: 48px;"></i>
        <p class="mt-3">Erreur lors du chargement des logs</p>
        <p class="small">${err.message}</p>
      </div>
    `;
  }
}

function setupListeners() {
  // Écouter les nouveaux logs en temps réel
  const onLogEntry = window.electronAPI?.onLogEntry
    ? (cb) => window.electronAPI.onLogEntry(cb)
    : window.api?.on
      ? (cb) => window.api.on('log-entry', cb)
      : null;

  if (onLogEntry) onLogEntry((logEntry) => {
    if (isPaused) return;
    if (!matchesFilters(logEntry)) return;

    liveQueue.push(logEntry);
    scheduleLiveFlush();
  });
}

function scheduleLiveFlush() {
  if (liveFlushTimer) return;
  liveFlushTimer = setTimeout(() => {
    liveFlushTimer = null;
    flushLiveQueue();
  }, 75);
}

function flushLiveQueue() {
  if (!liveQueue.length) return;

  const container = document.getElementById('logs-container');
  if (!container) {
    liveQueue = [];
    return;
  }

  const fragment = document.createDocumentFragment();
  const batch = liveQueue.splice(0, 200);
  for (const logEntry of batch) {
    const entry = createLogEntryElement(logEntry);
    fragment.appendChild(entry);
    stats[logEntry.level] = (stats[logEntry.level] || 0) + 1;
    totalLogs++;
  }
  container.appendChild(fragment);

  // Limiter le nombre d'entrées visibles pour éviter les ralentissements
  const maxVisible = 500;
  while (container.children.length > maxVisible) {
    container.removeChild(container.firstChild);
  }

  updateStats();
  if (autoScroll) scrollToBottom();

  if (liveQueue.length) {
    scheduleLiveFlush();
  }
}

function setupFilterListeners() {
  // Niveau
  document.getElementById('filter-level').addEventListener('change', (e) => {
    currentFilters.level = e.target.value;
  });
  
  // Catégorie
  document.getElementById('filter-category').addEventListener('change', (e) => {
    currentFilters.category = e.target.value;
  });
  
  // Recherche avec debounce
  let searchTimeout;
  document.getElementById('filter-search').addEventListener('input', (e) => {
    currentFilters.search = e.target.value;
    clearTimeout(searchTimeout);
    searchTimeout = setTimeout(() => applyFilters(), 300);
  });
  
  // Auto-scroll
  document.getElementById('auto-scroll').addEventListener('change', (e) => {
    autoScroll = e.target.checked;
  });
}

function matchesFilters(log) {
  // Filtre niveau
  if (currentFilters.level && currentFilters.level !== 'ALL') {
    const levelValues = { DEBUG: 0, INFO: 1, SUCCESS: 1, WARN: 2, ERROR: 3 };
    const filterLevel = levelValues[currentFilters.level];
    const logLevel = levelValues[log.level];
    if (logLevel < filterLevel) {
      return false;
    }
  }
  
  // Filtre catégorie
  if (currentFilters.category !== 'ALL' && log.category !== currentFilters.category) {
    return false;
  }
  
  // Filtre recherche
  if (currentFilters.search) {
    const searchLower = currentFilters.search.toLowerCase();
    const matches = log.message.toLowerCase().includes(searchLower) ||
                   (log.data && log.data.toLowerCase().includes(searchLower));
    if (!matches) return false;
  }
  
  return true;
}

function appendLogEntry(log, updateStat = true) {
  const container = document.getElementById('logs-container');
  const entry = createLogEntryElement(log);
  container.appendChild(entry);
  
  // Limiter le nombre d'entrées visibles pour éviter les ralentissements
  const maxVisible = 500;
  if (container.children.length > maxVisible) {
    container.removeChild(container.firstChild);
  }
}

function createLogEntryElement(log) {
  const entry = document.createElement('div');
  entry.className = `log-entry ${log.level}`;

  const time = new Date(log.timestamp).toLocaleTimeString('fr-FR', {
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit'
  });

  let html = `
    <span class="log-timestamp">${time}</span>
    <span class="log-category">${log.categoryLabel || log.category}</span>
    <span class="log-level" style="color: ${log.levelColor}">${log.levelEmoji} ${log.level}</span>
    <span class="log-message">${escapeHtml(log.message)}</span>
  `;

  if (log.data) {
    html += `<div class="log-data">${escapeHtml(log.data)}</div>`;
  }

  entry.innerHTML = html;
  return entry;
}

function escapeHtml(text) {
  if (!text) return '';
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

function updateStats() {
  document.getElementById('stat-debug').textContent = stats.DEBUG || 0;
  document.getElementById('stat-info').textContent = stats.INFO || 0;
  document.getElementById('stat-success').textContent = stats.SUCCESS || 0;
  document.getElementById('stat-warn').textContent = stats.WARN || 0;
  document.getElementById('stat-error').textContent = stats.ERROR || 0;
  document.getElementById('stat-total').textContent = totalLogs;
}

function scrollToBottom() {
  const container = document.getElementById('logs-container');
  container.scrollTop = container.scrollHeight;
}

async function applyFilters() {
  console.log('🔄 Application des filtres:', currentFilters);
  stats = { DEBUG: 0, INFO: 0, SUCCESS: 0, WARN: 0, ERROR: 0 };
  totalLogs = 0;
  await loadHistory();
}

async function clearLogs() {
  if (!confirm('⚠️ Êtes-vous sûr de vouloir effacer tous les logs ?\n\nCette action est irréversible.')) {
    return;
  }
  
  try {
    await invokeBridge('api-clear-logs');
    document.getElementById('logs-container').innerHTML = `
      <div class="text-center text-success py-5">
        <i class="bi bi-check-circle" style="font-size: 48px;"></i>
        <p class="mt-3">Logs effacés avec succès</p>
      </div>
    `;
    stats = { DEBUG: 0, INFO: 0, SUCCESS: 0, WARN: 0, ERROR: 0 };
    totalLogs = 0;
    updateStats();
    console.log('✅ Logs effacés');
  } catch (err) {
    console.error('❌ Erreur lors de l\'effacement:', err);
    alert('Erreur lors de l\'effacement des logs: ' + err.message);
  }
}

function pauseUpdates() {
  isPaused = !isPaused;
  const btn = document.getElementById('pause-text');
  if (isPaused) {
    btn.textContent = 'Reprendre';
    btn.parentElement.classList.remove('btn-outline-primary');
    btn.parentElement.classList.add('btn-warning');
    btn.previousElementSibling.className = 'bi bi-play-circle';
    console.log('⏸️ Mise en pause des mises à jour');
  } else {
    btn.textContent = 'Pause';
    btn.parentElement.classList.remove('btn-warning');
    btn.parentElement.classList.add('btn-outline-primary');
    btn.previousElementSibling.className = 'bi bi-pause-circle';
    console.log('▶️ Reprise des mises à jour');
  }
}

async function exportLogs() {
  try {
    const history = await invokeBridge('api-get-log-history', currentFilters);
    
    // Créer un fichier texte
    let content = '=== LOGS MAIL MONITOR ===\n';
    content += `Exporté le: ${new Date().toLocaleString('fr-FR')}\n`;
    content += `Filtres: Niveau=${currentFilters.level}, Catégorie=${currentFilters.category}\n`;
    content += `Total: ${history.length} entrées\n`;
    content += '='.repeat(80) + '\n\n';
    
    history.forEach(log => {
      const time = new Date(log.timestamp).toLocaleString('fr-FR');
      content += `[${time}] [${log.level}] [${log.category}]\n`;
      content += `  ${log.message}\n`;
      if (log.data) {
        content += `  Data: ${log.data}\n`;
      }
      content += '\n';
    });
    
    // Télécharger le fichier
    const blob = new Blob([content], { type: 'text/plain' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `logs-${new Date().toISOString().split('T')[0]}.txt`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    
    console.log('✅ Logs exportés');
  } catch (err) {
    console.error('❌ Erreur lors de l\'export:', err);
    alert('Erreur lors de l\'export: ' + err.message);
  }
}
