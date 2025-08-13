/**
 * Gestionnaire de l'affichage hiérarchique des dossiers monitorés
 */

class FoldersTreeManager {
  constructor(containerSelector = '#folders-tree') {
    this.container = document.querySelector(containerSelector);
    this.folders = new Map(); // Structure hiérarchique des dossiers
    this.filteredFolders = new Set();
    this.searchTerm = '';
    this.categoryFilter = '';
    
    this.init();
  }

  init() {
    this.setupEventListeners();
    this.loadFolders();
  }

  setupEventListeners() {
    // Recherche
    const searchInput = document.getElementById('folder-search');
    if (searchInput) {
      searchInput.addEventListener('input', (e) => {
        this.searchTerm = e.target.value.toLowerCase();
        this.filterFolders();
      });
    }

    // Filtre par catégorie
    const categoryFilter = document.getElementById('category-filter');
    if (categoryFilter) {
      categoryFilter.addEventListener('change', (e) => {
        this.categoryFilter = e.target.value;
        this.filterFolders();
      });
    }

    // Bouton actualiser
    const refreshBtn = document.getElementById('refresh-folders');
    if (refreshBtn) {
      refreshBtn.addEventListener('click', () => {
        this.loadFolders(true);
      });
    }

    // Bouton ajouter
    const addBtn = document.getElementById('add-folder');
  // Ne pas binder ici pour éviter les doubles ouvertures; l'app principale gère le clic

    // Toggle de vue (Tableau / Arborescence / Liste)
    const viewBoardBtn = document.getElementById('view-board');
    const viewTreeBtn = document.getElementById('view-tree');
    const viewListBtn = document.getElementById('view-list');
    const setActive = (btn) => {
      [viewBoardBtn, viewTreeBtn, viewListBtn].forEach(b => b?.classList.remove('active'));
      btn?.classList.add('active');
    };
    viewBoardBtn?.addEventListener('click', () => { setActive(viewBoardBtn); this.renderBoard(); });
    viewTreeBtn?.addEventListener('click', () => { setActive(viewTreeBtn); this.renderTree(); });
    viewListBtn?.addEventListener('click', () => { setActive(viewListBtn); this.renderList(); });

    // Déplier/Replier tout
    const expandAllBtn = document.getElementById('expand-all');
    const collapseAllBtn = document.getElementById('collapse-all');
    expandAllBtn?.addEventListener('click', () => {
      if (this.folders) this.expandAllNodes(this.folders);
      this.renderCurrentView();
    });
    collapseAllBtn?.addEventListener('click', () => {
      this.toggleAll(false);
      this.renderCurrentView();
    });
  }

  async loadFolders(forceRefresh = false) {
    try {
      console.log('📁 Chargement des dossiers monitorés...');
      this.showLoading();
      
  // Utiliser l'API IPC d'Electron
  const data = await window.electronAPI.invoke('api-folders-tree', { force: !!forceRefresh });
      console.log('📁 Données reçues de api-folders-tree:', data);

      if (!data || !data.folders) {
        console.warn('⚠️ Réponse invalide du serveur:', data);
        throw new Error('Réponse invalide du serveur');
      }

      console.log(`📁 ${data.folders.length} dossiers trouvés dans la réponse`);
      
  this.folders = this.buildFolderTree(data.folders || []);
  // Déplier tout par défaut
  this.expandAllNodes(this.folders);
  console.log('📁 Arbre construit, taille:', this.folders.size);
  this.renderCurrentView();
      this.updateStats(data.stats || {});

    } catch (error) {
      console.error('❌ Erreur chargement dossiers:', error);
      this.showError('Impossible de charger les dossiers: ' + error.message);
    }
  }

  buildFolderTree(folders) {
    console.log('🏗️ Construction arbre, données reçues:', folders);
    const tree = new Map();
    const pathMap = new Map();

    // Créer la structure hiérarchique complète à partir des chemins surveillés
    const allPaths = new Set();
    folders.forEach(folder => {
      const parts = folder.path.split('\\').filter(p => p.length > 0);
      let current = '';
      for (const part of parts) {
        current = current ? `${current}\\${part}` : part;
        allPaths.add(current);
      }
    });

    // Créer un map des dossiers pour retrouver les infos (isMonitored, category, emailCount)
    const folderInfoByPath = new Map();
    folders.forEach(folder => {
      folderInfoByPath.set(folder.path, folder);
    });

  // Construire l'arbre à partir de tous les chemins
    Array.from(allPaths).forEach(fullPath => {
      const pathParts = fullPath.split('\\').filter(part => part.length > 0);
      let currentPath = '';
      let parentNode = tree;
      pathParts.forEach((part, index) => {
        currentPath = currentPath ? `${currentPath}\\${part}` : part;
        if (!parentNode.has(part)) {
          const isLeaf = index === pathParts.length - 1;
          const info = folderInfoByPath.get(currentPath) || {};
          const node = {
            name: part,
            fullPath: currentPath,
            children: new Map(),
            isFolder: true,
            isMonitored: isLeaf && info.isMonitored,
            category: isLeaf ? info.category : null,
            emailCount: isLeaf ? info.emailCount : 0,
      isExpanded: true,
            level: index
          };
          parentNode.set(part, node);
          pathMap.set(currentPath, node);
        }
        parentNode = parentNode.get(part).children;
      });
    });

    console.log('🏗️ Arbre final construit, taille:', tree.size);
    console.log('🏗️ Contenu arbre:', Array.from(tree.keys()));
    return tree;
  }

  renderTree() {
  if (!this.container) return;
  // Ensure correct view visibility
  const boardEl = document.getElementById('folders-board');
  const treeEl = document.getElementById('folders-tree');
  if (boardEl && treeEl) { boardEl.classList.add('d-none'); treeEl.classList.remove('d-none'); }

    console.log('🎨 renderTree appelé, this.folders.size:', this.folders.size);
    console.log('🎨 Contenu de this.folders:', this.folders);
    // Préserver la position de scroll avant re-render
    let prevScrollTop = 0;
    const existingScrollEl = this.container.querySelector('.folders-tree');
    if (existingScrollEl) {
      prevScrollTop = existingScrollEl.scrollTop;
    } else {
      prevScrollTop = this.container.scrollTop || 0;
    }
    this.container.innerHTML = '';
    
    if (this.folders.size === 0) {
      console.log('🎨 Aucun dossier à afficher - message vide');
  this.container.innerHTML = `
        <div class="text-center text-muted py-4">
          <i class="bi bi-folder-x display-6 mb-3"></i>
          <p>Aucun dossier configuré</p>
  <button class="btn btn-outline-primary btn-sm" onclick="(window.app && window.app.showAddFolderModal) ? window.app.showAddFolderModal() : (window.foldersTree && window.foldersTree.showAddFolderModal && window.foldersTree.showAddFolderModal())">
            <i class="bi bi-plus me-1"></i>Ajouter un dossier
          </button>
        </div>
      `;
      return;
    }

    console.log('🎨 Affichage de l\'arbre avec', this.folders.size, 'nœuds racine');
    const treeElement = document.createElement('div');
    treeElement.className = 'folders-tree';
    
    this.renderNode(this.folders, treeElement, 0);
    this.container.appendChild(treeElement);
  // Tree-only view; no breadcrumbs in board layout
    // Restaurer la position de scroll après re-render
    const newScrollEl = this.container.querySelector('.folders-tree');
    if (newScrollEl) {
      newScrollEl.scrollTop = prevScrollTop;
    } else {
      this.container.scrollTop = prevScrollTop;
    }
  }

  renderBoard() {
    const container = document.getElementById('monitoring-content');
    if (!container) return;
    // Toggle visibility
    const boardEl = document.getElementById('folders-board');
    const treeEl = document.getElementById('folders-tree');
    if (boardEl && treeEl) { boardEl.classList.remove('d-none'); treeEl.classList.add('d-none'); }
    // Build board columns by category
    const board = boardEl;
    board.innerHTML = '';
    const cols = [
      { key: 'Déclarations', icon: '📋', title: 'Déclarations' },
      { key: 'Règlements', icon: '💰', title: 'Règlements' },
      { key: 'Mails simples', icon: '📧', title: 'Mails simples' }
    ];
    const itemsByCat = new Map();
    const includeNode = (n) => (this.filteredFolders.size === 0 || this.filteredFolders.has(n.fullPath));
    const walk = (map) => map.forEach(n => {
      if (n.isMonitored && includeNode(n)) {
        const k = n.category || 'Mails simples';
        if (!itemsByCat.has(k)) itemsByCat.set(k, []);
        itemsByCat.get(k).push(n);
      }
      if (n.children?.size) walk(n.children);
    });
    if (this.folders && this.folders.size) walk(this.folders);
    cols.forEach(c => {
      const col = document.createElement('div');
      col.className = 'board-col';
      col.innerHTML = `
        <div class="board-header">
          <div class="fw-semibold">${c.icon} ${c.title}</div>
          <span class="badge bg-light text-dark">${(itemsByCat.get(c.key) || []).length}</span>
        </div>
        <div class="board-body" data-cat="${c.key}"></div>
      `;
      board.appendChild(col);
      const body = col.querySelector('.board-body');
      (itemsByCat.get(c.key) || []).forEach(n => {
        const item = document.createElement('div');
        item.className = 'board-item';
        item.dataset.path = n.fullPath;
        item.innerHTML = `
          <div>
            <div class="fw-semibold">${this.escapeHtml(n.name)}</div>
            <div class="meta">${this.escapeHtml(n.fullPath)}</div>
          </div>
          <div class="actions d-flex align-items-center gap-1">
            <span class="badge bg-light text-dark">${n.emailCount || 0}</span>
            <button class="btn btn-outline-secondary btn-sm" data-act="details" title="Détails"><i class="bi bi-eye"></i></button>
            <button class="btn btn-outline-primary btn-sm" data-act="edit" title="Modifier"><i class="bi bi-pencil"></i></button>
            <button class="btn btn-outline-danger btn-sm" data-act="remove" title="Retirer"><i class="bi bi-trash"></i></button>
          </div>
        `;
        body.appendChild(item);
      });
    });
    // Bind actions for board items
    board.querySelectorAll('[data-act="details"]').forEach(btn => btn.addEventListener('click', (e) => {
      const el = e.currentTarget.closest('[data-path]');
      const path = el?.getAttribute('data-path');
      const node = this.findNodeByPath(path);
      if (node) this.showDetailsModal(node);
    }));
    board.querySelectorAll('[data-act="edit"]').forEach(btn => btn.addEventListener('click', (e) => {
      const el = e.currentTarget.closest('[data-path]');
      const path = el?.getAttribute('data-path');
      const node = this.findNodeByPath(path);
      if (node) this.editFolder(node);
    }));
    board.querySelectorAll('[data-act="remove"]').forEach(btn => btn.addEventListener('click', (e) => {
      const el = e.currentTarget.closest('[data-path]');
      const path = el?.getAttribute('data-path');
      const node = this.findNodeByPath(path);
      if (node) this.removeFromMonitoring(node);
    }));
  }

  renderList() {
  const treeEl = document.getElementById('folders-tree');
  const boardEl = document.getElementById('folders-board');
  if (!treeEl || !boardEl) return;
  boardEl.classList.add('d-none');
  treeEl.classList.remove('d-none');
  const prevScrollTop = treeEl.scrollTop || 0;
  treeEl.innerHTML = '';
  const list = document.createElement('div');
  list.className = 'folders-list';
    const items = [];
    const walk = (map) => map.forEach(node => { if (this.filteredFolders.size === 0 || this.filteredFolders.has(node.fullPath)) { if (node.isMonitored) items.push(node); } if (node.children?.size) walk(node.children); });
    if (this.folders && this.folders.size) walk(this.folders);
    if (items.length === 0) {
      list.innerHTML = `<div class="text-muted text-center py-4">Aucun dossier monitoré</div>`;
    } else {
      list.innerHTML = items.map(n => `
        <div class="d-flex align-items-center justify-content-between border rounded p-2 mb-2" data-path="${n.fullPath}">
          <div>
            <div class="fw-semibold">${this.escapeHtml(n.name)}</div>
            <div class="text-muted small">${this.escapeHtml(n.fullPath)}</div>
          </div>
          <div class="d-flex align-items-center gap-2">
            ${n.category ? `<span class="badge bg-primary">${n.category}</span>` : ''}
            <span class="badge bg-light text-dark">${n.emailCount || 0}</span>
            <button class="btn btn-outline-primary btn-sm" data-act="edit">Modifier</button>
            <button class="btn btn-outline-danger btn-sm" data-act="remove">Retirer</button>
          </div>
        </div>
      `).join('');
    }
  treeEl.appendChild(list);
  treeEl.scrollTop = prevScrollTop;
    // Bind actions
    this.container.querySelectorAll('[data-act="edit"]').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const el = e.currentTarget.closest('[data-path]');
        const path = el?.getAttribute('data-path');
        const node = this.findNodeByPath(path);
        if (node) this.editFolder(node);
      });
    });
    this.container.querySelectorAll('[data-act="remove"]').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const el = e.currentTarget.closest('[data-path]');
        const path = el?.getAttribute('data-path');
        const node = this.findNodeByPath(path);
        if (node) this.removeFromMonitoring(node);
      });
    });
  // Board layout: no breadcrumbs
  }

  renderCurrentView() {
  const boardBtn = document.getElementById('view-board');
  const treeBtn = document.getElementById('view-tree');
  const listBtn = document.getElementById('view-list');
  if (boardBtn?.classList.contains('active')) this.renderBoard();
  else if (treeBtn?.classList.contains('active')) this.renderTree();
  else if (listBtn?.classList.contains('active')) this.renderList();
  }

  // Déplie récursivement tous les nœuds
  expandAllNodes(nodeMap) {
    nodeMap.forEach((node) => {
      node.isExpanded = true;
      if (node.children && node.children.size > 0) {
        this.expandAllNodes(node.children);
      }
    });
  }

  renderNode(nodeMap, parentElement, level) {
    nodeMap.forEach((node, key) => {
      if (this.filteredFolders.size > 0 && !this.filteredFolders.has(node.fullPath)) {
        return; // Filtrage actif et ce nœud ne correspond pas
      }

      const nodeElement = document.createElement('div');
      nodeElement.className = 'folder-node';
      nodeElement.dataset.path = node.fullPath;
      nodeElement.dataset.level = level;

      const hasChildren = node.children.size > 0;
      const isExpanded = node.isExpanded;
      const isLeafNode = !hasChildren; // Seuls les nœuds feuilles peuvent être configurés

      // Infos additionnelles (exemple: chemin complet, date dernier email, non lus)
      let extraInfo = '';
      if (node.fullPath) {
        extraInfo += `<div class='folder-extra'><span class='text-muted'>Chemin:</span> <span class='folder-path'>${node.fullPath}</span></div>`;
      }
      if (node.lastEmailDate) {
        extraInfo += `<div class='folder-extra'><span class='text-muted'>Dernier email:</span> <span>${node.lastEmailDate}</span></div>`;
      }
      if (typeof node.unreadCount === 'number') {
        extraInfo += `<div class='folder-extra'><span class='text-muted'>Non lus:</span> <span>${node.unreadCount}</span></div>`;
      }

      nodeElement.innerHTML = `
        <div class="folder-item modern-card ${node.isMonitored ? 'monitored' : ''}" data-path="${node.fullPath}" style="padding-left: ${level * 18}px;">
          <div class="folder-connector"></div>
          <div class="folder-icon ${hasChildren ? 'expandable' : ''}">
            ${hasChildren ? 
              `<i class="bi bi-chevron-right ${isExpanded ? 'expanded' : ''}"></i>` :
              `<i class="bi bi-folder${node.isMonitored ? '-check' : ''}"></i>`
            }
          </div>
          <div class="folder-name ${level === 0 ? 'root' : ''}" title="${node.fullPath}">
            ${this.highlightSearchTerm(node.name)}
          </div>
          ${extraInfo}
          ${node.category ? `
            <span class="category-badge category-${node.category.toLowerCase().replace(/\s+/g, '-')}">
              ${node.category}
            </span>
          ` : ''}
          ${node.emailCount > 0 ? `
            <span class="badge bg-gradient-modern ms-1">${node.emailCount}</span>
          ` : ''}
          ${node.isMonitored ? `
            <div class="monitoring-indicator active" title="Dossier surveillé"></div>
          ` : ''}
          <div class="folder-actions modern-actions">
            ${node.isMonitored ? `
              <button class="btn btn-modern edit-folder" title="Modifier">
                <i class="bi bi-pencil"></i>
              </button>
              <button class="btn btn-modern remove-folder" title="Supprimer du monitoring">
                <i class="bi bi-trash"></i>
              </button>
            ` : `
              <button class="btn btn-modern add-to-monitoring" title="Ajouter au monitoring">
                <i class="bi bi-plus"></i>
              </button>
            `}
          </div>
        </div>
      `;

      // Gestionnaires d'événements pour ce nœud
      this.setupNodeEventListeners(nodeElement, node);

      parentElement.appendChild(nodeElement);

      // Render children
      if (hasChildren) {
        const childrenContainer = document.createElement('div');
        childrenContainer.className = `folder-children ${isExpanded ? '' : 'collapsed'}`;
        this.renderNode(node.children, childrenContainer, level + 1);
        nodeElement.appendChild(childrenContainer);
      }
    });
  }

  setupNodeEventListeners(nodeElement, node) {
    // Expansion/collapse
    const expandIcon = nodeElement.querySelector('.folder-icon.expandable');
    if (expandIcon) {
      expandIcon.addEventListener('click', (e) => {
        e.stopPropagation();
        this.toggleNode(node);
      });
    }

    // Actions des boutons
    const editBtn = nodeElement.querySelector('.edit-folder');
    if (editBtn) {
      editBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        this.editFolder(node);
      });
    }

    const removeBtn = nodeElement.querySelector('.remove-folder');
    if (removeBtn) {
      removeBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        this.removeFromMonitoring(node);
      });
    }

    const addBtn = nodeElement.querySelector('.add-to-monitoring');
    if (addBtn) {
      addBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        this.addToMonitoring(node);
      });
    }

    // Sélection du dossier
    const folderItem = nodeElement.querySelector('.folder-item');
    folderItem.addEventListener('click', () => {
      this.selectFolder(node);
    });
  }

  toggleNode(node) {
    node.isExpanded = !node.isExpanded;
  this.renderCurrentView(); // Re-render pour montrer/cacher les enfants
  }

  selectFolder(node) {
    // Retirer la sélection précédente
    document.querySelectorAll('.folder-item.selected').forEach(el => {
      el.classList.remove('selected');
    });

    // Sélectionner le nouveau dossier
  const host = this.getElementByPath(node.fullPath);
  const folderElement = host?.querySelector('.folder-item') || host;
  if (folderElement) folderElement.classList.add('selected');

    // Mettre à jour le panneau de détails
    this.renderDetailsPanel(node);
    // Émettre un événement pour d'autres composants
    document.dispatchEvent(new CustomEvent('folderSelected', { detail: { folder: node } }));
  }

  renderDetailsPanel(node) {
    this.showDetailsModal(node);
  }

  showDetailsModal(node) {
    const modalEl = document.getElementById('folderDetailsModal');
    const titleEl = document.getElementById('folderDetailTitle');
    const bodyEl = document.getElementById('folderDetailBody');
    const footerEl = document.getElementById('folderDetailFooter');
    if (!modalEl || !titleEl || !bodyEl || !footerEl) return;
    const categoryBadge = node.category ? `<span class="badge bg-primary">${node.category}</span>` : '<span class="badge bg-secondary">Non catégorisé</span>';
    const unread = typeof node.unreadCount === 'number' ? node.unreadCount : 0;
    const emailCount = typeof node.emailCount === 'number' ? node.emailCount : 0;
    titleEl.textContent = node.name || 'Détails du dossier';
    bodyEl.innerHTML = `
      <div class="mb-2">${categoryBadge}</div>
      <div class="text-muted small mb-3">${this.escapeHtml(node.fullPath)}</div>
      <div class="row g-2">
        <div class="col-6"><div class="border rounded p-2 text-center"><div class="small text-muted">Total</div><div class="fw-bold">${emailCount}</div></div></div>
        <div class="col-6"><div class="border rounded p-2 text-center"><div class="small text-muted">Non lus</div><div class="fw-bold">${unread}</div></div></div>
      </div>
    `;
    footerEl.innerHTML = node.isMonitored ? `
      <button class="btn btn-outline-primary" id="modal-edit">Modifier</button>
      <button class="btn btn-outline-danger" id="modal-remove">Retirer</button>
      <button class="btn btn-secondary" data-bs-dismiss="modal">Fermer</button>
    ` : `
      <button class="btn btn-primary" id="modal-add">Ajouter au monitoring</button>
      <button class="btn btn-secondary" data-bs-dismiss="modal">Fermer</button>
    `;
    footerEl.querySelector('#modal-edit')?.addEventListener('click', () => this.editFolder(node));
    footerEl.querySelector('#modal-remove')?.addEventListener('click', () => this.removeFromMonitoring(node));
    footerEl.querySelector('#modal-add')?.addEventListener('click', () => this.addToMonitoring(node));
    try { new bootstrap.Modal(modalEl).show(); } catch (_) {}
  }

  escapeHtml(text) {
    if (!text) return '';
    const map = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;' };
    return String(text).replace(/[&<>"']/g, m => map[m]);
  }

  async editFolder(node) {
  // Close details modal if open before editing
  this.hideModal('folderDetailsModal');
    const selected = await this.showCategoryModal('Modifier la catégorie', node.category || 'Déclarations', 'Enregistrer');
    if (!selected) return;
    await this.updateFolderCategory(node, selected);
  }

  async updateFolderCategory(node, newCategory) {
    try {
  let folderElement = this.getElementByPath(node.fullPath);
      if (folderElement) {
        folderElement.classList.add('updating');
      }

      const result = await window.electronAPI.invoke('api-folders-update-category', {
        folderPath: node.fullPath,
        category: newCategory
      });

      if (!result || !result.success) {
        throw new Error(result?.error || 'Erreur lors de la mise à jour');
      }

      node.category = newCategory;
      
      if (folderElement) {
        folderElement.classList.remove('updating');
        folderElement.classList.add('success');
        setTimeout(() => folderElement.classList.remove('success'), 2000);
      }

  // Force reload from backend to avoid any stale cache and ensure persistence
  await this.loadFolders(true);
      this.showSuccess(`Catégorie mise à jour: ${newCategory}`);

    } catch (error) {
      console.error('❌ Erreur mise à jour catégorie:', error);
      const folderElement = this.getElementByPath(node.fullPath);
      if (folderElement) {
        folderElement.classList.remove('updating');
        folderElement.classList.add('error');
        setTimeout(() => folderElement.classList.remove('error'), 3000);
      }
      this.showError('Erreur lors de la mise à jour');
    }
  }

  async removeFromMonitoring(node) {
  const ok = await this.showConfirmModal('Retirer du monitoring', `Supprimer "${this.escapeHtml(node.name)}" du monitoring ?`, 'Retirer');
  if (!ok) return;
  // Close details modal if open before applying removal
  this.hideModal('folderDetailsModal');

    try {
      const result = await window.electronAPI.invoke('api-folders-remove', {
        folderPath: node.fullPath
      });

      if (!result || !result.success) {
        throw new Error(result?.error || 'Erreur lors de la suppression');
      }

      node.isMonitored = false;
      node.category = null;
      
  // Force reload from backend so deletion is reflected and stats updated
  await this.loadFolders(true);
      this.showSuccess(`Dossier retiré du monitoring`);

    } catch (error) {
      console.error('❌ Erreur suppression:', error);
      this.showError('Erreur lors de la suppression');
    }
  }

  async addToMonitoring(node) {
  // Close details modal if open before adding
  this.hideModal('folderDetailsModal');
    const selected = await this.showCategoryModal('Ajouter au monitoring', 'Déclarations', 'Ajouter');
    if (!selected) return;
    await this.addFolderToMonitoring(node.fullPath, selected);
  }

  // Category selection modal used by Edit/Add
  showCategoryModal(title, initialCategory = 'Déclarations', confirmText = 'Valider') {
    return new Promise((resolve) => {
      const id = 'categoryModalFolders';
      const existing = document.getElementById(id);
      if (existing) existing.remove();
      const html = `
        <div class="modal fade" id="${id}" tabindex="-1" aria-hidden="true">
          <div class="modal-dialog">
            <div class="modal-content">
              <div class="modal-header">
                <h6 class="modal-title">${title}</h6>
                <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
              </div>
              <div class="modal-body">
                <label class="form-label">Catégorie</label>
                <select class="form-select" id="cat-select">
                  <option value="Déclarations" ${initialCategory === 'Déclarations' ? 'selected' : ''}>Déclarations</option>
                  <option value="Règlements" ${initialCategory === 'Règlements' ? 'selected' : ''}>Règlements</option>
                  <option value="Mails simples" ${initialCategory === 'Mails simples' ? 'selected' : ''}>Mails simples</option>
                </select>
              </div>
              <div class="modal-footer">
                <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Annuler</button>
                <button type="button" class="btn btn-primary" id="cat-ok">${confirmText}</button>
              </div>
            </div>
          </div>
        </div>`;
      document.body.insertAdjacentHTML('beforeend', html);
      const modalEl = document.getElementById(id);
      const modal = new bootstrap.Modal(modalEl);
      const cleanup = (val) => { try { modal.hide(); } catch(_) {} setTimeout(() => modalEl.remove(), 200); resolve(val); };
      modalEl.querySelector('#cat-ok').addEventListener('click', () => {
        const val = modalEl.querySelector('#cat-select').value;
        cleanup(val || null);
      });
      modalEl.addEventListener('hidden.bs.modal', () => cleanup(null), { once: true });
      modal.show();
    });
  }

  // Robust selector for [data-path] elements (paths contain backslashes)
  getElementByPath(path) {
    const all = document.querySelectorAll('[data-path]');
    for (const el of all) {
      if (el.getAttribute('data-path') === path) return el;
    }
    return null;
  }

  // Pretty confirmation modal using Bootstrap
  showConfirmModal(title, message, confirmText = 'Confirmer') {
    return new Promise((resolve) => {
      const id = 'confirmModalFolders';
      const existing = document.getElementById(id);
      if (existing) existing.remove();
      const html = `
        <div class="modal fade" id="${id}" tabindex="-1" aria-hidden="true">
          <div class="modal-dialog modal-sm">
            <div class="modal-content">
              <div class="modal-header">
                <h6 class="modal-title">${title}</h6>
                <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
              </div>
              <div class="modal-body"><div class="small">${message}</div></div>
              <div class="modal-footer">
                <button type="button" class="btn btn-secondary btn-sm" data-bs-dismiss="modal">Annuler</button>
                <button type="button" class="btn btn-danger btn-sm" id="confirm-ok">${confirmText}</button>
              </div>
            </div>
          </div>
        </div>`;
      document.body.insertAdjacentHTML('beforeend', html);
      const modalEl = document.getElementById(id);
      const modal = new bootstrap.Modal(modalEl);
      const cleanup = (result) => { try { modal.hide(); } catch(_) {} setTimeout(() => modalEl.remove(), 200); resolve(result); };
      modalEl.querySelector('#confirm-ok').addEventListener('click', () => cleanup(true));
      modalEl.addEventListener('hidden.bs.modal', () => cleanup(false), { once: true });
      modal.show();
    });
  }

  // Helper: hide a Bootstrap modal by id if visible
  hideModal(id) {
    const el = document.getElementById(id);
    if (!el) return;
    try {
      const instance = bootstrap.Modal.getInstance(el) || new bootstrap.Modal(el);
      instance.hide();
    } catch (_) { /* noop */ }
  }

  filterFolders() {
    this.filteredFolders.clear();

    if (!this.searchTerm && !this.categoryFilter) {
      this.renderCurrentView();
      return;
    }

    // Filtrer récursivement
    this.filterNodeRecursive(this.folders);
    this.renderCurrentView();
  }

  filterNodeRecursive(nodeMap) {
    nodeMap.forEach((node, key) => {
      let matches = false;

      // Test de recherche textuelle
      if (this.searchTerm && node.name.toLowerCase().includes(this.searchTerm)) {
        matches = true;
      }

      // Test de filtre par catégorie
      if (this.categoryFilter && node.category === this.categoryFilter) {
        matches = true;
      }

      // Test récursif des enfants
      if (node.children.size > 0) {
        this.filterNodeRecursive(node.children);
        // Si un enfant correspond, inclure le parent
        node.children.forEach((child) => {
          if (this.filteredFolders.has(child.fullPath)) {
            matches = true;
          }
        });
      }

      if (matches) {
        this.filteredFolders.add(node.fullPath);
      }
    });
  }

  highlightSearchTerm(text) {
    if (!this.searchTerm) return text;
    
    const regex = new RegExp(`(${this.searchTerm})`, 'gi');
    return text.replace(regex, '<span class="folder-search-highlight">$1</span>');
  }

  updateStats(stats) {
  const totalEl = document.getElementById('total-folders');
  if (totalEl) totalEl.textContent = stats.total || 0;
  const activeEl = document.getElementById('active-folders');
  if (activeEl) activeEl.textContent = stats.active || 0;
  const declEl = document.getElementById('declarations-count');
  if (declEl) declEl.textContent = stats.declarations || 0;
  const reglEl = document.getElementById('reglements-count');
  if (reglEl) reglEl.textContent = stats.reglements || 0;
  const simplesEl = document.getElementById('simples-count');
  if (simplesEl) simplesEl.textContent = stats.simples || 0;
  }

  showLoading() {
    if (this.container) {
      this.container.innerHTML = `
        <div class="text-center text-muted py-4">
          <div class="spinner-border spinner-border-sm me-2" role="status"></div>
          Chargement des dossiers...
        </div>
      `;
    }
  }

  showError(message) {
    if (this.container) {
      this.container.innerHTML = `
        <div class="alert alert-danger">
          <i class="bi bi-exclamation-triangle me-2"></i>
          ${message}
        </div>
      `;
    }
  }

  showSuccess(message) {
    // Utiliser le système de notification existant
    if (window.mailMonitor && window.mailMonitor.showNotification) {
      window.mailMonitor.showNotification('Succès', message, 'success');
    }
  }

  showAddFolderModal() {
    // Déléguer au sélecteur moderne de l'application (COM rapide + EWS fallback)
    if (window.app && typeof window.app.showAddFolderModal === 'function') {
      try {
        window.app.showAddFolderModal();
        return;
      } catch (_) { /* fallback below */ }
    }

    // Fallback: modal manuel basique si le flux principal n'est pas dispo
    const id = 'legacyManualAddFolderModal';
    const existing = document.getElementById(id);
    if (existing) existing.remove();
    const html = `
      <div class="modal fade" id="${id}" tabindex="-1">
        <div class="modal-dialog">
          <div class="modal-content">
            <div class="modal-header">
              <h6 class="modal-title">Ajouter un dossier (manuel)</h6>
              <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
            </div>
            <div class="modal-body">
              <div class="mb-2">
                <label class="form-label">Chemin complet</label>
                <input type="text" class="form-control" id="legacy-folder-path" placeholder="Ex: Boîte de réception\\Déclarations" required />
              </div>
              <div class="mb-2">
                <label class="form-label">Catégorie</label>
                <select class="form-select" id="legacy-folder-cat" required>
                  <option value="">-- Sélectionner une catégorie --</option>
                  <option value="Déclarations">Déclarations</option>
                  <option value="Règlements">Règlements</option>
                  <option value="Mails simples">Mails simples</option>
                </select>
              </div>
            </div>
            <div class="modal-footer">
              <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Annuler</button>
              <button type="button" class="btn btn-primary" id="legacy-folder-save">Ajouter</button>
            </div>
          </div>
        </div>
      </div>`;
    document.body.insertAdjacentHTML('beforeend', html);
    const modalEl = document.getElementById(id);
    const modal = new bootstrap.Modal(modalEl);
    modal.show();
    modalEl.querySelector('#legacy-folder-save').addEventListener('click', async () => {
      const path = document.getElementById('legacy-folder-path').value.trim();
      const category = document.getElementById('legacy-folder-cat').value.trim();
      if (!path || !category) return;
      try {
        await this.addFolderToMonitoring(path, category);
        try { modal.hide(); } catch(_) {}
        setTimeout(() => modalEl.remove(), 200);
      } catch (_) { /* errors handled in addFolderToMonitoring */ }
    });
  }

  async addFolderToMonitoring(folderPath, category) {
    try {
      const result = await window.electronAPI.invoke('api-folders-add', {
        folderPath: folderPath,
        category: category
      });

      if (!result || !result.success) {
        throw new Error(result?.error || 'Erreur lors de l\'ajout');
      }

  this.loadFolders(true); // Recharger la liste
  const count = result.count || 1;
  this.showSuccess(`${count} dossier(s) ajouté(s) au monitoring`);

    } catch (error) {
      console.error('❌ Erreur ajout dossier:', error);
      this.showError('Erreur lors de l\'ajout du dossier');
    }
  }

  // Utilitaires: recherche et fil d'Ariane
  findNodeByPath(path) {
    if (!path || !this.folders) return null;
    let found = null;
    const traverse = (map) => {
      for (const node of map.values()) {
        if (node.fullPath === path) { found = node; return; }
        if (node.children && node.children.size > 0) traverse(node.children);
        if (found) return;
      }
    };
    traverse(this.folders);
    return found;
  }

  updateBreadcrumbs(node) { /* Breadcrumbs removed in new design */ }

  toggleAll(expand) {
    const walk = (map) => map.forEach(n => { n.isExpanded = !!expand; if (n.children?.size) walk(n.children); });
    if (this.folders) walk(this.folders);
  }
}

// Initialiser le gestionnaire de dossiers quand le DOM est prêt
// Expose l'instance sur window pour les handlers inline (ex: bouton "Ajouter un dossier" en état vide)
document.addEventListener('DOMContentLoaded', () => {
  window.foldersTree = new FoldersTreeManager();
});

// Export pour utilisation dans d'autres modules
if (typeof module !== 'undefined' && module.exports) {
  module.exports = FoldersTreeManager;
}
