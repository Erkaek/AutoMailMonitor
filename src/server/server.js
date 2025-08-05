const express = require('express');
const path = require('path');
const outlookConnector = require('./outlookConnector');

const app = express();
const PORT = 3000;

// Configuration middleware
app.use(express.json());
app.use(express.static(path.join(__dirname, '../../public')));

// Headers CORS et sécurité
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  res.header('X-Content-Type-Options', 'nosniff');
  res.header('X-Frame-Options', 'DENY');
  next();
});

// === ROUTES API ===

// Status Outlook
app.get('/api/outlook/status', async (req, res) => {
  try {
    const status = await outlookConnector.isConnected();
    const connectionInfo = outlookConnector.getConnectionInfo();
    
    res.json({ 
      status,
      timestamp: new Date().toISOString(),
      version: process.env.npm_package_version || '1.0.0',
      details: connectionInfo
    });
  } catch (error) {
    console.error('Erreur statut Outlook:', error);
    res.status(500).json({ 
      error: 'Impossible de vérifier le statut Outlook',
      details: error.message 
    });
  }
});

// Statistiques générales
app.get('/api/stats/summary', async (req, res) => {
  try {
    // Récupérer les statistiques depuis la base de données
    const databaseService = require('./services/databaseService');
    
    // Statistiques d'aujourd'hui depuis la base de données
    const today = new Date().toISOString().split('T')[0];
    const emailsToday = await databaseService.getEmailCountByDate(today);
    const sentToday = await databaseService.getSentEmailCountByDate(today);
    const totalEmails = await databaseService.getTotalEmailCount();
    const unreadCount = await databaseService.getUnreadEmailCount();
    
    // Calculer le temps de réponse moyen depuis la base de données
    const avgResponseTime = await databaseService.getAverageResponseTime();
    
    const stats = {
      emailsToday: emailsToday || 0,
      sentToday: sentToday || 0,
      unreadTotal: unreadCount || 0,
      avgResponseTime: avgResponseTime || "0.0",
      lastSync: new Date().toISOString(),
      outlookConnected: await outlookConnector.isConnected(),
      message: "Données depuis la base de données locale",
      folders: {
        inbox: totalEmails || 0,
        sent: sentToday || 0,
        drafts: 0,
        junk: 0
      }
    };
    
    res.json(stats);
  } catch (error) {
    console.error('Erreur stats:', error);
    res.status(500).json({ 
      error: 'Impossible de récupérer les statistiques',
      details: error.message 
    });
  }
});

// Emails récents
app.get('/api/emails/recent', async (req, res) => {
  try {
    const limit = parseInt(req.query.limit) || 10;
    const offset = parseInt(req.query.offset) || 0;
    
    // Récupérer les emails depuis la base de données
    const databaseService = require('./services/databaseService');
    const emails = await databaseService.getRecentEmails(limit);
    const totalCount = await databaseService.getTotalEmailCount();
    
    res.json({
      emails: emails || [],
      total: totalCount || 0,
      limit,
      offset,
      hasMore: offset + limit < (totalCount || 0)
    });
  } catch (error) {
    console.error('Erreur emails récents:', error);
    res.status(500).json({ 
      error: 'Impossible de récupérer les emails',
      details: error.message 
    });
  }
});

// Analytics données
app.get('/api/analytics/daily', async (req, res) => {
  try {
    const days = parseInt(req.query.days) || 7;
    const analytics = generateDailyAnalytics(days);
    
    res.json(analytics);
  } catch (error) {
    console.error('Erreur analytics:', error);
    res.status(500).json({ 
      error: 'Impossible de récupérer les analytics',
      details: error.message 
    });
  }
});

// Analytics par dossier
app.get('/api/analytics/folders', async (req, res) => {
  try {
    const folderStats = {
      inbox: { count: 156, percentage: 62.4 },
      sent: { count: 89, percentage: 35.6 },
      drafts: { count: 3, percentage: 1.2 },
      junk: { count: 2, percentage: 0.8 }
    };
    
    res.json(folderStats);
  } catch (error) {
    console.error('Erreur analytics dossiers:', error);
    res.status(500).json({ 
      error: 'Impossible de récupérer les analytics de dossiers',
      details: error.message 
    });
  }
});

// Configuration utilisateur
app.get('/api/settings', (req, res) => {
  // Retourne les paramètres par défaut avec configuration des dossiers
  const defaultSettings = {
    syncInterval: 60,
    autoStart: false,
    minimizeToTray: true,
    notificationsEnabled: true,
    watchInbox: true,
    watchSent: false,
    emailLimit: 50,
    retentionDays: 30,
    folderCategories: {
      // Configuration des dossiers avec leurs catégories
      'Inbox': { category: 'mails_simple', enabled: true },
      'Sent': { category: 'mails_simple', enabled: false },
      'Déclarations': { category: 'declaration', enabled: true },
      'Règlements': { category: 'reglement', enabled: true },
      'Factures': { category: 'reglement', enabled: true },
      'Commandes': { category: 'declaration', enabled: true }
    },
    availableCategories: [
      { id: 'declaration', name: 'Déclarations', color: '#0d6efd', icon: 'file-text' },
      { id: 'reglement', name: 'Règlements', color: '#198754', icon: 'credit-card' },
      { id: 'mails_simple', name: 'Mails simples', color: '#6c757d', icon: 'envelope' }
    ]
  };
  
  res.json(defaultSettings);
});

// Récupération des dossiers Outlook disponibles
app.get('/api/outlook/folders', async (req, res) => {
  try {
    const isConnected = await outlookConnector.isConnected();
    
    if (!isConnected) {
      return res.json({
        folders: [],
        message: 'Outlook non connecté. Démarrez Outlook pour voir les dossiers disponibles.'
      });
    }
    
    const folders = outlookConnector.getFolders();
    res.json({
      folders: folders.map(folder => ({
        name: folder.name,
        path: folder.path,
        type: folder.type,
        itemCount: folder.itemCount || 0,
        unreadCount: folder.unreadCount || 0
      })),
      total: folders.length,
      lastUpdate: new Date().toISOString()
    });
  } catch (error) {
    console.error('Erreur récupération dossiers:', error);
    res.status(500).json({ 
      error: 'Impossible de récupérer les dossiers Outlook',
      details: error.message 
    });
  }
});

// Récupération des boîtes mail connectées
app.get('/api/outlook/mailboxes', async (req, res) => {
  try {
    const isConnected = await outlookConnector.isConnected();
    
    if (!isConnected) {
      return res.json({
        mailboxes: [],
        message: 'Outlook non connecté. Démarrez Outlook pour voir les boîtes mail disponibles.'
      });
    }
    
    const mailboxes = await outlookConnector.getMailboxes();
    res.json({
      mailboxes,
      total: mailboxes.length,
      lastUpdate: new Date().toISOString()
    });
  } catch (error) {
    console.error('Erreur récupération boîtes mail:', error);
    res.status(500).json({ 
      error: 'Impossible de récupérer les boîtes mail',
      details: error.message 
    });
  }
});

// Récupération de la structure complète des dossiers d'une boîte mail
app.get('/api/outlook/folder-structure/:storeId?', async (req, res) => {
  try {
    const isConnected = await outlookConnector.isConnected();
    
    if (!isConnected) {
      return res.json({
        structure: null,
        message: 'Outlook non connecté'
      });
    }
    
    const { storeId } = req.params;
    const structure = await outlookConnector.getFolderStructure(storeId);
    
    res.json({
      structure,
      storeId: storeId || 'default',
      lastUpdate: new Date().toISOString()
    });
  } catch (error) {
    console.error('Erreur structure dossiers:', error);
    res.status(500).json({ 
      error: 'Impossible de récupérer la structure des dossiers',
      details: error.message 
    });
  }
});

// Configuration des catégories de dossiers
app.post('/api/settings/folders', (req, res) => {
  try {
    const { folderCategories } = req.body;
    
    if (!folderCategories || typeof folderCategories !== 'object') {
      return res.status(400).json({ 
        error: 'Configuration de dossiers invalide' 
      });
    }
    
    // Validation des catégories
    const validCategories = ['declaration', 'reglement', 'mails_simple'];
    
    for (const [folderPath, config] of Object.entries(folderCategories)) {
      if (!validCategories.includes(config.category)) {
        return res.status(400).json({ 
          error: `Catégorie invalide: ${config.category}` 
        });
      }
    }
    
    // Ici on sauvegarderait dans une vraie base de données
    console.log('Configuration dossiers mise à jour:', folderCategories);
    
    res.json({ 
      success: true, 
      message: 'Configuration des dossiers sauvegardée',
      folderCategories 
    });
  } catch (error) {
    console.error('Erreur sauvegarde dossiers:', error);
    res.status(500).json({ 
      error: 'Impossible de sauvegarder la configuration des dossiers',
      details: error.message 
    });
  }
});

// Statistiques par catégorie
app.get('/api/stats/by-category', async (req, res) => {
  try {
    const period = req.query.period || 'today';
    
    // Simulation de données par catégorie
    // Dans une vraie implémentation, on croiserait les données Outlook avec la config des dossiers
    const statsByCategory = {
      declaration: {
        emailsReceived: Math.floor(Math.random() * 20) + 5,
        emailsSent: Math.floor(Math.random() * 10) + 2,
        unreadCount: Math.floor(Math.random() * 15) + 3,
        folders: ['Déclarations', 'Commandes']
      },
      reglement: {
        emailsReceived: Math.floor(Math.random() * 15) + 3,
        emailsSent: Math.floor(Math.random() * 8) + 1,
        unreadCount: Math.floor(Math.random() * 10) + 1,
        folders: ['Règlements', 'Factures']
      },
      mails_simple: {
        emailsReceived: Math.floor(Math.random() * 30) + 10,
        emailsSent: Math.floor(Math.random() * 15) + 5,
        unreadCount: Math.floor(Math.random() * 25) + 8,
        folders: ['Inbox', 'Sent']
      }
    };
    
    res.json({
      period,
      categories: statsByCategory,
      lastUpdate: new Date().toISOString()
    });
  } catch (error) {
    console.error('Erreur stats par catégorie:', error);
    res.status(500).json({ 
      error: 'Impossible de récupérer les statistiques par catégorie',
      details: error.message 
    });
  }
});

app.post('/api/settings', (req, res) => {
  try {
    const settings = req.body;
    
    // Validation basique
    if (typeof settings.syncInterval !== 'number' || settings.syncInterval < 10) {
      return res.status(400).json({ 
        error: 'Intervalle de synchronisation invalide (minimum 10 secondes)' 
      });
    }
    
    // Ici on sauvegarderait dans une vraie base de données
    console.log('Nouveaux paramètres reçus:', settings);
    
    res.json({ 
      success: true, 
      message: 'Paramètres sauvegardés',
      settings 
    });
  } catch (error) {
    console.error('Erreur sauvegarde paramètres:', error);
    res.status(500).json({ 
      error: 'Impossible de sauvegarder les paramètres',
      details: error.message 
    });
  }
});

// Action: Forcer la synchronisation
app.post('/api/sync/force', async (req, res) => {
  try {
    console.log('Synchronisation forcée demandée');
    
    // Simulation du processus de sync
    await new Promise(resolve => setTimeout(resolve, 2000));
    
    res.json({ 
      success: true, 
      message: 'Synchronisation terminée',
      timestamp: new Date().toISOString(),
      emailsProcessed: Math.floor(Math.random() * 20) + 5
    });
  } catch (error) {
    console.error('Erreur sync forcée:', error);
    res.status(500).json({ 
      error: 'Échec de la synchronisation',
      details: error.message 
    });
  }
});

// === ROUTES GESTION DES DOSSIERS ===

// Récupérer l'arbre hiérarchique des dossiers
app.get('/api/folders/tree', async (req, res) => {
  try {
    const databaseService = require('../services/databaseService');
    await databaseService.initialize();

    // Récupérer les dossiers configurés
    const foldersConfig = await databaseService.getFoldersConfiguration();
    
    // Récupérer la structure Outlook pour affichage hiérarchique
    const allFolders = await outlookConnector.getFolderStructure();
    
    // Combiner avec les configurations
    const enrichedFolders = [];
    
    for (const [folderPath, config] of Object.entries(foldersConfig)) {
      // Chercher le dossier dans la structure Outlook
      const outlookFolder = findFolderInStructure(allFolders, folderPath);
      
      enrichedFolders.push({
        path: folderPath,
        name: extractFolderName(folderPath),
        isMonitored: true,
        category: config.category || 'Mails simples',
        emailCount: outlookFolder ? outlookFolder.Count || 0 : 0,
        parentPath: getParentPath(folderPath)
      });
    }

    // Ajouter les dossiers non configurés pour affichage
    addNonConfiguredFolders(allFolders, enrichedFolders, foldersConfig);

    // Calculer les statistiques
    const stats = calculateFolderStats(enrichedFolders);

    res.json({
      folders: enrichedFolders,
      stats: stats,
      timestamp: new Date().toISOString()
    });

  } catch (error) {
    console.error('❌ Erreur récupération arbre dossiers:', error);
    res.status(500).json({
      error: 'Impossible de récupérer les dossiers',
      details: error.message
    });
  }
});

// Ajouter un dossier au monitoring
app.post('/api/folders/add', async (req, res) => {
  try {
    const { folderPath, category } = req.body;
    
    if (!folderPath || !category) {
      return res.status(400).json({
        error: 'Chemin du dossier et catégorie requis'
      });
    }

    const databaseService = require('../services/databaseService');
    await databaseService.initialize();

    // Vérifier que le dossier existe dans Outlook
    const folderExists = await outlookConnector.folderExists(folderPath);
    if (!folderExists) {
      return res.status(404).json({
        error: 'Dossier non trouvé dans Outlook'
      });
    }

    // Sauvegarder la configuration
    const currentConfig = await databaseService.getFoldersConfiguration();
    currentConfig[folderPath] = {
      category: category,
      name: extractFolderName(folderPath),
      isActive: true
    };

    await databaseService.saveFoldersConfiguration(currentConfig);

    res.json({
      success: true,
      message: 'Dossier ajouté au monitoring',
      folderPath: folderPath,
      category: category
    });

  } catch (error) {
    console.error('❌ Erreur ajout dossier:', error);
    res.status(500).json({
      error: 'Impossible d\'ajouter le dossier',
      details: error.message
    });
  }
});

// Mettre à jour la catégorie d'un dossier
app.post('/api/folders/update-category', async (req, res) => {
  try {
    const { folderPath, category } = req.body;
    
    if (!folderPath || !category) {
      return res.status(400).json({
        error: 'Chemin du dossier et catégorie requis'
      });
    }

    const databaseService = require('../services/databaseService');
    await databaseService.initialize();

    const currentConfig = await databaseService.getFoldersConfiguration();
    
    if (!currentConfig[folderPath]) {
      return res.status(404).json({
        error: 'Dossier non trouvé dans la configuration'
      });
    }

    currentConfig[folderPath].category = category;
    await databaseService.saveFoldersConfiguration(currentConfig);

    res.json({
      success: true,
      message: 'Catégorie mise à jour',
      folderPath: folderPath,
      category: category
    });

  } catch (error) {
    console.error('❌ Erreur mise à jour catégorie:', error);
    res.status(500).json({
      error: 'Impossible de mettre à jour la catégorie',
      details: error.message
    });
  }
});

// Supprimer un dossier du monitoring
app.delete('/api/folders/remove', async (req, res) => {
  try {
    const { folderPath } = req.body;
    
    if (!folderPath) {
      return res.status(400).json({
        error: 'Chemin du dossier requis'
      });
    }

    const databaseService = require('../services/databaseService');
    await databaseService.initialize();

    const currentConfig = await databaseService.getFoldersConfiguration();
    
    if (!currentConfig[folderPath]) {
      return res.status(404).json({
        error: 'Dossier non trouvé dans la configuration'
      });
    }

    delete currentConfig[folderPath];
    await databaseService.saveFoldersConfiguration(currentConfig);

    res.json({
      success: true,
      message: 'Dossier retiré du monitoring',
      folderPath: folderPath
    });

  } catch (error) {
    console.error('❌ Erreur suppression dossier:', error);
    res.status(500).json({
      error: 'Impossible de supprimer le dossier',
      details: error.message
    });
  }
});

// === FONCTIONS UTILITAIRES POUR LES DOSSIERS ===

function findFolderInStructure(folders, targetPath) {
  for (const folder of folders) {
    if (folder.FolderPath === targetPath) {
      return folder;
    }
    if (folder.SubFolders && folder.SubFolders.length > 0) {
      const found = findFolderInStructure(folder.SubFolders, targetPath);
      if (found) return found;
    }
  }
  return null;
}

function extractFolderName(folderPath) {
  const parts = folderPath.split('\\');
  return parts[parts.length - 1] || folderPath;
}

function getParentPath(folderPath) {
  const parts = folderPath.split('\\');
  if (parts.length <= 1) return null;
  return parts.slice(0, -1).join('\\');
}

function addNonConfiguredFolders(outlookFolders, enrichedFolders, foldersConfig, parentPath = '') {
  outlookFolders.forEach(folder => {
    const fullPath = parentPath ? `${parentPath}\\${folder.Name}` : folder.Name;
    
    // Ajouter seulement si pas déjà configuré
    if (!foldersConfig[fullPath]) {
      enrichedFolders.push({
        path: fullPath,
        name: folder.Name,
        isMonitored: false,
        category: null,
        emailCount: folder.Count || 0,
        parentPath: parentPath || null
      });
    }

    // Traiter récursivement les sous-dossiers
    if (folder.SubFolders && folder.SubFolders.length > 0) {
      addNonConfiguredFolders(folder.SubFolders, enrichedFolders, foldersConfig, fullPath);
    }
  });
}

function calculateFolderStats(folders) {
  const stats = {
    total: folders.length,
    active: 0,
    declarations: 0,
    reglements: 0,
    simples: 0
  };

  folders.forEach(folder => {
    if (folder.isMonitored) {
      stats.active++;
      
      switch (folder.category) {
        case 'Déclarations':
          stats.declarations++;
          break;
        case 'Règlements':
          stats.reglements++;
          break;
        case 'Mails simples':
          stats.simples++;
          break;
      }
    }
  });

  return stats;
}

// Health check
app.get('/api/health', (req, res) => {
  res.json({
    status: 'ok',
    timestamp: new Date().toISOString(),
    uptime: process.uptime(),
    memory: process.memoryUsage(),
    version: process.env.npm_package_version || '1.0.0'
  });
});

// === FONCTIONS UTILITAIRES ===

// Démarrage du serveur
async function setupServer() {
  console.log('🔧 Configuration du serveur Express...');
  
  return new Promise((resolve, reject) => {
    const server = app.listen(PORT, () => {
      console.log(`🚀 Mail Monitor Server démarré sur http://localhost:${PORT}`);
      console.log(`📧 API disponible sur http://localhost:${PORT}/api`);
      console.log(`⚡ Environnement: ${process.env.NODE_ENV || 'development'}`);
      resolve(server);
    });

    server.on('error', (error) => {
      console.error('❌ Erreur serveur:', error);
      reject(error);
    });
  });
}

// Version synchrone pour compatibilité
setupServer.sync = function() {
  const subjects = [
    'Rapport mensuel - Janvier 2025',
    'Réunion équipe - Planning hebdomadaire',
    'Nouvelle commande #CMD-2025-001',
    'Facture #INV-2025-001 - À traiter',
    'Rappel: Formation sécurité obligatoire',
    'Mise à jour système - Maintenance programmée',
    'Présentation projet Q1 2025',
    'Demande de congés - Validation requise',
    'Newsletter entreprise - Janvier',
    'Invitation: Séminaire innovation',
    'Rapport d\'incident #INC-2025-003',
    'Proposition commerciale - Client ABC',
    'Confirmation rendez-vous médical',
    'Mise à jour politique RGPD',
    'Félicitations équipe ventes!'
  ];
  
  const senders = [
    { name: 'Marie Dupont', email: 'marie.dupont@entreprise.com' },
    { name: 'Service Comptabilité', email: 'compta@entreprise.com' },
    { name: 'Plateforme E-commerce', email: 'no-reply@platform.com' },
    { name: 'Jean Martin', email: 'jean.martin@partenaire.org' },
    { name: 'Système de notifications', email: 'notifications@systeme.net' },
    { name: 'RH - Ressources Humaines', email: 'rh@entreprise.com' },
    { name: 'Support Client', email: 'support@service.fr' },
    { name: 'Direction Générale', email: 'direction@entreprise.com' }
  ];

  const emails = [];
  
  for (let i = 0; i < limit; i++) {
    const index = offset + i;
    const sender = senders[index % senders.length];
    const subject = subjects[index % subjects.length];
    
    // Heure aléatoire dans les dernières 48h
    const hoursAgo = Math.random() * 48;
    const receivedTime = new Date(Date.now() - hoursAgo * 60 * 60 * 1000);
    
    emails.push({
      id: `email_${index + 1}`,
      subject: `${subject} ${index > subjects.length ? '#' + (index - subjects.length) : ''}`,
      sender: sender.name,
      senderEmail: sender.email,
      preview: generateEmailPreview(subject),
      receivedTime: receivedTime.toISOString(),
      isUnread: Math.random() > 0.7,
      hasAttachment: Math.random() > 0.8,
      importance: Math.random() > 0.9 ? 'high' : 'normal',
      size: Math.floor(Math.random() * 500) + 50 // KB
    });
  }
  
  return emails;
}

function generateEmailPreview(subject) {
  const previews = {
    'Rapport mensuel': 'Veuillez trouver ci-joint le rapport mensuel avec les indicateurs...',
    'Réunion équipe': 'La réunion hebdomadaire aura lieu jeudi prochain à 14h en salle...',
    'Nouvelle commande': 'Une nouvelle commande vient d\'être passée sur notre plateforme...',
    'Facture': 'Veuillez trouver ci-joint la facture pour vos services du mois...',
    'Formation': 'Rappel: la formation sécurité obligatoire se déroulera le...',
    'Maintenance': 'Une maintenance système est programmée dimanche prochain de...',
    'Présentation': 'Merci de préparer la présentation pour le projet Q1 avec...',
    'Congés': 'Votre demande de congés pour la période du... est en attente...',
    'Newsletter': 'Découvrez les actualités de notre entreprise ce mois-ci...',
    'Séminaire': 'Vous êtes cordialement invité(e) au séminaire sur l\'innovation...'
  };
  
  for (const [key, preview] of Object.entries(previews)) {
    if (subject.includes(key)) {
      return preview;
    }
  }
  
  return 'Aperçu du contenu de l\'email non disponible...';
}

function generateDailyAnalytics(days) {
  const analytics = {
    period: {
      start: new Date(Date.now() - (days - 1) * 24 * 60 * 60 * 1000).toISOString().split('T')[0],
      end: new Date().toISOString().split('T')[0],
      days
    },
    daily: [],
    summary: {
      totalReceived: 0,
      totalSent: 0,
      avgDaily: 0,
      peakDay: null
    }
  };
  
  let maxEmails = 0;
  let peakDay = null;
  
  for (let i = days - 1; i >= 0; i--) {
    const date = new Date(Date.now() - i * 24 * 60 * 60 * 1000);
    const received = Math.floor(Math.random() * 40) + 10;
    const sent = Math.floor(Math.random() * 20) + 5;
    
    if (received > maxEmails) {
      maxEmails = received;
      peakDay = date.toISOString().split('T')[0];
    }
    
    analytics.daily.push({
      date: date.toISOString().split('T')[0],
      received,
      sent,
      total: received + sent
    });
    
    analytics.summary.totalReceived += received;
    analytics.summary.totalSent += sent;
  }
  
  analytics.summary.avgDaily = Math.round((analytics.summary.totalReceived + analytics.summary.totalSent) / days);
  analytics.summary.peakDay = peakDay;
  
  return analytics;
}

// === GESTION D'ERREURS ===

// 404 pour routes API non trouvées
app.use('/api/*', (req, res) => {
  res.status(404).json({ 
    error: 'Endpoint non trouvé',
    path: req.path,
    method: req.method 
  });
});

// Route par défaut pour l'interface
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, '../../public/index.html'));
});

// Gestionnaire d'erreurs global
app.use((error, req, res, next) => {
  console.error('Erreur serveur:', error);
  res.status(500).json({ 
    error: 'Erreur interne du serveur',
    message: process.env.NODE_ENV === 'development' ? error.message : 'Une erreur est survenue'
  });
});

// === DÉMARRAGE SERVEUR ===

function setupServer() {
  return new Promise((resolve, reject) => {
    try {
      const server = app.listen(PORT, () => {
        console.log(`🚀 Mail Monitor Server démarré sur http://localhost:${PORT}`);
        console.log(`📧 API disponible sur http://localhost:${PORT}/api`);
        console.log(`⚡ Environnement: ${process.env.NODE_ENV || 'development'}`);
        
        // Résoudre la promesse une fois le serveur démarré
        resolve(server);
      });

      server.on('error', (error) => {
        console.error('❌ Erreur serveur:', error);
        reject(error);
      });

      // Gestion propre de l'arrêt
      process.on('SIGTERM', () => {
        console.log('🛑 Arrêt du serveur demandé...');
        server.close(() => {
          console.log('✅ Serveur arrêté proprement');
          process.exit(0);
        });
      });

      process.on('SIGINT', () => {
        console.log('\n🛑 Interruption détectée (Ctrl+C)');
        server.close(() => {
          console.log('✅ Serveur arrêté proprement');
          process.exit(0);
        });
      });

    } catch (error) {
      console.error('❌ Erreur démarrage serveur:', error);
      reject(error);
    }
  });
}

// Version synchrone pour compatibilité
setupServer.sync = function() {
  const server = app.listen(PORT, () => {
    console.log(`🚀 Mail Monitor Server démarré sur http://localhost:${PORT}`);
    console.log(`📧 API disponible sur http://localhost:${PORT}/api`);
    console.log(`⚡ Environnement: ${process.env.NODE_ENV || 'development'}`);
  });

  server.on('error', (error) => {
    console.error('❌ Erreur serveur:', error);
  });

  return server;
};

module.exports = setupServer;
