/**
 * Service de synchronisation intelligente entre Outlook et la base de données
 * Remplace le système d'événements par une synchronisation directe
 */

class EmailSyncService {
  constructor(databaseService, outlookConnector) {
    this.db = databaseService;
    this.outlook = outlookConnector;
    this.lastSyncTime = null;
  }

  /**
   * Synchronise un email spécifique en comparant son état actuel avec Outlook
   */
  async syncEmail(emailId) {
    try {
      // Récupérer l'email depuis la base de données
      const dbEmail = await this.db.getEmailById(emailId);
      if (!dbEmail) {
        console.log(`⚠️ Email ${emailId} non trouvé en base`);
        return false;
      }

      // Récupérer l'email depuis Outlook
      const outlookEmail = await this.outlook.getEmailById(dbEmail.outlook_id);
      if (!outlookEmail) {
        console.log(`⚠️ Email ${dbEmail.outlook_id} non trouvé dans Outlook`);
        return false;
      }

      // Comparer les états et détecter les changements
      const changes = this.detectChanges(dbEmail, outlookEmail);
      
      if (changes.length > 0) {
        console.log(`🔄 Synchronisation email ${emailId}: ${changes.length} changement(s)`);
        
        // Appliquer les changements en base
        await this.applyChanges(emailId, changes);
        
        changes.forEach(change => {
          console.log(`   ${change.field}: ${change.oldValue} → ${change.newValue}`);
        });
        
        return true;
      }

      return false; // Aucun changement
    } catch (error) {
      console.error(`❌ Erreur sync email ${emailId}:`, error);
      return false;
    }
  }

  /**
   * Détecte les différences entre l'email en base et celui d'Outlook
   */
  detectChanges(dbEmail, outlookEmail) {
    const changes = [];

    // Vérifier le statut de lecture
    if (dbEmail.is_read !== outlookEmail.UnRead) {
      changes.push({
        field: 'is_read',
        oldValue: dbEmail.is_read,
        newValue: !outlookEmail.UnRead // Inverser car UnRead = true signifie non lu
      });
    }

    // Vérifier le statut de traitement (si l'email a été déplacé)
    const isProcessed = this.isEmailProcessed(outlookEmail);
    if (dbEmail.is_processed !== isProcessed) {
      changes.push({
        field: 'is_processed',
        oldValue: dbEmail.is_processed,
        newValue: isProcessed
      });
    }

    // Vérifier la catégorie (si elle a changé)
    const currentCategory = this.determineEmailCategory(outlookEmail);
    if (dbEmail.category !== currentCategory) {
      changes.push({
        field: 'category',
        oldValue: dbEmail.category,
        newValue: currentCategory
      });
    }

    // Vérifier l'importance
    if (dbEmail.importance !== outlookEmail.Importance) {
      changes.push({
        field: 'importance',
        oldValue: dbEmail.importance,
        newValue: outlookEmail.Importance
      });
    }

    return changes;
  }

  /**
   * Applique les changements détectés en base de données
   */
  async applyChanges(emailId, changes) {
    try {
      const updateFields = {};
      
      changes.forEach(change => {
        updateFields[change.field] = change.newValue;
      });

      // Ajouter la date de dernière mise à jour
      updateFields.updated_at = new Date().toISOString();

      await this.db.updateEmail(emailId, updateFields);
      
      console.log(`✅ Email ${emailId} mis à jour avec ${changes.length} changement(s)`);
    } catch (error) {
      console.error(`❌ Erreur application changements email ${emailId}:`, error);
      throw error;
    }
  }

  /**
   * Synchronise tous les emails de la base avec Outlook
   */
  async syncAllEmails() {
    try {
      console.log('🔄 Début de la synchronisation complète...');
      
      const allEmails = await this.db.getAllEmails();
      console.log(`📧 ${allEmails.length} emails à synchroniser`);

      let syncedCount = 0;
      let changedCount = 0;

      for (const email of allEmails) {
        const hasChanged = await this.syncEmail(email.id);
        syncedCount++;
        
        if (hasChanged) {
          changedCount++;
        }

        // Petite pause pour ne pas surcharger Outlook
        if (syncedCount % 10 === 0) {
          await this.sleep(100);
          console.log(`   📊 Progression: ${syncedCount}/${allEmails.length} emails traités`);
        }
      }

      this.lastSyncTime = new Date();
      
      console.log(`✅ Synchronisation terminée:`);
      console.log(`   📧 ${syncedCount} emails synchronisés`);
      console.log(`   🔄 ${changedCount} emails mis à jour`);
      console.log(`   ⏰ Dernière sync: ${this.lastSyncTime.toLocaleString()}`);

      return {
        total: syncedCount,
        changed: changedCount,
        lastSync: this.lastSyncTime
      };

    } catch (error) {
      console.error('❌ Erreur synchronisation complète:', error);
      throw error;
    }
  }

  /**
   * Synchronisation incrémentale (seulement les emails récents)
   */
  async syncRecentEmails(hoursBack = 24) {
    try {
      const cutoffDate = new Date();
      cutoffDate.setHours(cutoffDate.getHours() - hoursBack);

      console.log(`🔄 Synchronisation incrémentale (${hoursBack}h)...`);
      
      const recentEmails = await this.db.getEmailsSince(cutoffDate);
      console.log(`📧 ${recentEmails.length} emails récents à synchroniser`);

      let changedCount = 0;

      for (const email of recentEmails) {
        const hasChanged = await this.syncEmail(email.id);
        if (hasChanged) {
          changedCount++;
        }
      }

      console.log(`✅ Sync incrémentale terminée: ${changedCount}/${recentEmails.length} emails mis à jour`);
      
      return changedCount;

    } catch (error) {
      console.error('❌ Erreur synchronisation incrémentale:', error);
      throw error;
    }
  }

  /**
   * Détermine si un email a été traité (déplacé vers un dossier spécifique)
   */
  isEmailProcessed(outlookEmail) {
    try {
      const folderName = outlookEmail.Parent?.Name || '';
      
      // Considérer comme traité si dans certains dossiers
      const processedFolders = ['Traités', 'Processed', 'Archive', 'Completed'];
      return processedFolders.some(folder => 
        folderName.toLowerCase().includes(folder.toLowerCase())
      );
    } catch (error) {
      return false;
    }
  }

  /**
   * Détermine la catégorie d'un email basée sur des règles
   */
  determineEmailCategory(outlookEmail) {
    try {
      const subject = outlookEmail.Subject || '';
      const sender = outlookEmail.SenderEmailAddress || '';
      
      // Règles de catégorisation
      if (subject.includes('Déclaration') || subject.includes('TVA')) {
        return 'Déclarations';
      }
      
      if (subject.includes('Facture') || subject.includes('Paiement')) {
        return 'Règlements';
      }
      
      return 'Mails simples';
    } catch (error) {
      return 'Mails simples';
    }
  }

  /**
   * Utilitaire pour les pauses
   */
  sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  /**
   * Démarre une synchronisation automatique périodique
   */
  startAutoSync(intervalMinutes = 15) {
    console.log(`🔄 Démarrage auto-sync (toutes les ${intervalMinutes} minutes)`);
    
    this.autoSyncInterval = setInterval(async () => {
      try {
        console.log('🔄 Auto-sync en cours...');
        await this.syncRecentEmails(1); // Sync de la dernière heure
      } catch (error) {
        console.error('❌ Erreur auto-sync:', error);
      }
    }, intervalMinutes * 60 * 1000);
  }

  /**
   * Arrête la synchronisation automatique
   */
  stopAutoSync() {
    if (this.autoSyncInterval) {
      clearInterval(this.autoSyncInterval);
      this.autoSyncInterval = null;
      console.log('⏹️ Auto-sync arrêtée');
    }
  }
}

module.exports = EmailSyncService;
