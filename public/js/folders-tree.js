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
    if (addBtn) {
      addBtn.addEventListener('click', () => {
        this.showAddFolderModal();
      });
    }
  }

  async loadFolders(forceRefresh = false) {
    try {
      console.log('📁 Chargement des dossiers monitorés...');
      this.showLoading();
      
      // Utiliser l'API IPC d'Electron
      const data = await window.electronAPI.invoke('api-folders-tree');
      console.log('📁 Données reçues de api-folders-tree:', data);

      if (!data || !data.folders) {
        console.warn('⚠️ Réponse invalide du serveur:', data);
        throw new Error('Réponse invalide du serveur');
      }

      console.log(`📁 ${data.folders.length} dossiers trouvés dans la réponse`);
      
      this.folders = this.buildFolderTree(data.folders || []);
      console.log('📁 Arbre construit, taille:', this.folders.size);
      this.renderTree();
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

    // Créer la structure hiérarchique
    folders.forEach(folder => {
      console.log('📂 Traitement dossier:', folder);
      const pathParts = folder.path.split('\\').filter(part => part.length > 0);
      console.log('📂 Parties du chemin:', pathParts);
      let currentPath = '';
      let parentNode = tree;

      pathParts.forEach((part, index) => {
        currentPath = currentPath ? `${currentPath}\\${part}` : part;
        console.log(`📂 Partie ${index}: ${part}, chemin actuel: ${currentPath}`);
        
        if (!parentNode.has(part)) {
          const isLeaf = index === pathParts.length - 1;
          const node = {
            name: part,
            fullPath: currentPath,
            children: new Map(),
            isFolder: true,
            isMonitored: isLeaf && folder.isMonitored,
            category: isLeaf ? folder.category : null,
            emailCount: isLeaf ? folder.emailCount : 0,
            isExpanded: false,
            level: index
          };
          
          console.log(`📂 Création nœud: ${part}, isLeaf: ${isLeaf}, isMonitored: ${node.isMonitored}`);
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

    console.log('🎨 renderTree appelé, this.folders.size:', this.folders.size);
    console.log('🎨 Contenu de this.folders:', this.folders);

    this.container.innerHTML = '';
    
    if (this.folders.size === 0) {
      console.log('🎨 Aucun dossier à afficher - message vide');
    this.container.innerHTML = `
        <div class="text-center text-muted py-4">
          <i class="bi bi-folder-x display-6 mb-3"></i>
          <p>Aucun dossier configuré</p>
      <button class="btn btn-outline-primary btn-sm" onclick="window.foldersTree && window.foldersTree.showAddFolderModal()">
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

      nodeElement.innerHTML = `
        <div class="folder-item ${node.isMonitored ? 'monitored' : ''}" data-path="${node.fullPath}">
          <div class="folder-icon ${hasChildren ? 'expandable' : ''}">
            ${hasChildren ? 
              `<i class="bi bi-chevron-right ${isExpanded ? 'expanded' : ''}"></i>` :
              `<i class="bi bi-folder${node.isMonitored ? '-check' : ''}"></i>`
            }
          </div>
          <div class="folder-name ${level === 0 ? 'root' : ''}" title="${node.fullPath}">
            ${this.highlightSearchTerm(node.name)}
          </div>
          ${node.category ? `
            <span class="category-badge category-${node.category.toLowerCase().replace(/\s+/g, '-')}">
              ${node.category}
            </span>
          ` : ''}
          ${node.emailCount > 0 ? `
            <span class="badge bg-secondary ms-1">${node.emailCount}</span>
          ` : ''}
          ${node.isMonitored ? `
            <div class="monitoring-indicator active" title="Dossier surveillé"></div>
          ` : ''}
          ${isLeafNode ? `
            <div class="folder-actions">
              ${node.isMonitored ? `
                <button class="btn btn-outline-primary btn-sm edit-folder" title="Modifier">
                  <i class="bi bi-pencil"></i>
                </button>
                <button class="btn btn-outline-danger btn-sm remove-folder" title="Supprimer du monitoring">
                  <i class="bi bi-trash"></i>
                </button>
              ` : `
                <button class="btn btn-outline-success btn-sm add-to-monitoring" title="Ajouter au monitoring">
                  <i class="bi bi-plus"></i>
                </button>
              `}
            </div>
          ` : ''}
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
    this.renderTree(); // Re-render pour montrer/cacher les enfants
  }

  selectFolder(node) {
    // Retirer la sélection précédente
    document.querySelectorAll('.folder-item.selected').forEach(el => {
      el.classList.remove('selected');
    });

    // Sélectionner le nouveau dossier
    const folderElement = document.querySelector(`[data-path="${node.fullPath}"] .folder-item`);
    if (folderElement) {
      folderElement.classList.add('selected');
    }

    // Émettre un événement pour d'autres composants
    document.dispatchEvent(new CustomEvent('folderSelected', {
      detail: { folder: node }
    }));
  }

  async editFolder(node) {
    const folderElement = document.querySelector(`[data-path="${node.fullPath}"]`);
    if (!folderElement) return;

    // Créer le formulaire d'édition
    const editForm = document.createElement('div');
    editForm.className = 'folder-edit-form';
    editForm.innerHTML = `
      <div class="row g-2">
        <div class="col-md-6">
          <label class="form-label">Catégorie</label>
          <select class="form-select" id="edit-category">
            <option value="Déclarations" ${node.category === 'Déclarations' ? 'selected' : ''}>Déclarations</option>
            <option value="Règlements" ${node.category === 'Règlements' ? 'selected' : ''}>Règlements</option>
            <option value="Mails simples" ${node.category === 'Mails simples' ? 'selected' : ''}>Mails simples</option>
          </select>
        </div>
        <div class="col-md-6">
          <label class="form-label">Actions</label>
          <div class="d-flex gap-1">
            <button class="btn btn-success btn-sm save-changes">
              <i class="bi bi-check"></i> Sauver
            </button>
            <button class="btn btn-secondary btn-sm cancel-edit">
              <i class="bi bi-x"></i> Annuler
            </button>
          </div>
        </div>
      </div>
    `;

    // Insérer après l'élément du dossier
    folderElement.parentNode.insertBefore(editForm, folderElement.nextSibling);
    
    // Gestionnaires pour le formulaire
    editForm.querySelector('.save-changes').addEventListener('click', async () => {
      const newCategory = editForm.querySelector('#edit-category').value;
      await this.updateFolderCategory(node, newCategory);
      editForm.remove();
    });

    editForm.querySelector('.cancel-edit').addEventListener('click', () => {
      editForm.remove();
    });
  }

  async updateFolderCategory(node, newCategory) {
    try {
      folderElement = document.querySelector(`[data-path="${node.fullPath}"]`);
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

      this.renderTree();
      this.showSuccess(`Catégorie mise à jour: ${newCategory}`);

    } catch (error) {
      console.error('❌ Erreur mise à jour catégorie:', error);
      if (folderElement) {
        folderElement.classList.remove('updating');
        folderElement.classList.add('error');
        setTimeout(() => folderElement.classList.remove('error'), 3000);
      }
      this.showError('Erreur lors de la mise à jour');
    }
  }

  async removeFromMonitoring(node) {
    if (!confirm(`Supprimer "${node.name}" du monitoring ?`)) return;

    try {
      const result = await window.electronAPI.invoke('api-folders-remove', {
        folderPath: node.fullPath
      });

      if (!result || !result.success) {
        throw new Error(result?.error || 'Erreur lors de la suppression');
      }

      node.isMonitored = false;
      node.category = null;
      
      this.renderTree();
      this.showSuccess(`Dossier retiré du monitoring`);

    } catch (error) {
      console.error('❌ Erreur suppression:', error);
      this.showError('Erreur lors de la suppression');
    }
  }

  async addToMonitoring(node) {
  const category = prompt('Choisir une catégorie:\n1. Déclarations\n2. Règlements\n3. Mails simples', '1');
  const categories = { '1': 'Déclarations', '2': 'Règlements', '3': 'Mails simples' };
  const selectedCategory = categories[category];
  if (!selectedCategory) return;
  // Utiliser l'IPC Electron (serveur Express supprimé)
  await this.addFolderToMonitoring(node.fullPath, selectedCategory);
  }

  filterFolders() {
    this.filteredFolders.clear();

    if (!this.searchTerm && !this.categoryFilter) {
      this.renderTree();
      return;
    }

    // Filtrer récursivement
    this.filterNodeRecursive(this.folders);
    this.renderTree();
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
    document.getElementById('total-folders').textContent = stats.total || 0;
    document.getElementById('active-folders').textContent = stats.active || 0;
    document.getElementById('declarations-count').textContent = stats.declarations || 0;
    document.getElementById('reglements-count').textContent = stats.reglements || 0;
    document.getElementById('simples-count').textContent = stats.simples || 0;
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
    // Pour l'instant, utiliser prompt - peut être amélioré avec un modal Bootstrap
    const folderPath = prompt('Chemin du dossier Outlook à ajouter:');
    if (!folderPath) return;

    const category = prompt('Catégorie:\n1. Déclarations\n2. Règlements\n3. Mails simples', '1');
    const categories = {
      '1': 'Déclarations',
      '2': 'Règlements',
      '3': 'Mails simples'
    };

    const selectedCategory = categories[category];
    if (!selectedCategory) return;

    this.addFolderToMonitoring(folderPath, selectedCategory);
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

  this.loadFolders(); // Recharger la liste
  const count = result.count || 1;
  this.showSuccess(`${count} dossier(s) ajouté(s) au monitoring`);

    } catch (error) {
      console.error('❌ Erreur ajout dossier:', error);
      this.showError('Erreur lors de l\'ajout du dossier');
    }
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
