/**
 * Test script pour vérifier que les recipient_email sont bien récupérés et stockés
 */

const path = require('path');
const dbService = require('./src/services/optimizedDatabaseService');

async function testRecipientEmails() {
    console.log('🔍 Test de récupération des recipient_email');
    
    try {
        console.log('✅ Service de base de données chargé');
        
        // Initialiser le service si nécessaire
        if (!dbService.db) {
            await dbService.initialize();
            console.log('✅ Service initialisé');
        }
        
        // Récupérer les 10 derniers emails directement avec SQL
        const recentEmails = dbService.db.prepare(`
            SELECT * FROM emails 
            ORDER BY created_at DESC 
            LIMIT ?
        `).all(10);
        
        console.log(`📧 ${recentEmails.length} emails récents trouvés\n`);
        
        let emailsWithRecipients = 0;
        let emailsWithoutRecipients = 0;
        
        recentEmails.forEach((email, index) => {
            console.log(`${index + 1}. ${email.subject}`);
            console.log(`   Expéditeur: ${email.sender_name} <${email.sender_email}>`);
            
            if (email.recipient_email && email.recipient_email.trim() !== '') {
                console.log(`   ✅ Destinataires: ${email.recipient_email}`);
                emailsWithRecipients++;
            } else {
                console.log(`   ❌ Aucun destinataire enregistré`);
                emailsWithoutRecipients++;
            }
            
            console.log(`   Dossier: ${email.folder_name}`);
            console.log(`   Date: ${email.received_time}`);
            console.log('');
        });
        
        console.log('📊 Résumé:');
        console.log(`   Emails avec destinataires: ${emailsWithRecipients}`);
        console.log(`   Emails sans destinataires: ${emailsWithoutRecipients}`);
        
        if (emailsWithoutRecipients > 0) {
            console.log('\n⚠️  Des emails n\'ont pas de destinataires enregistrés.');
            console.log('   Cela peut être normal pour:');
            console.log('   - Les emails dans des dossiers système (Brouillons, Éléments supprimés)');
            console.log('   - Les anciens emails ajoutés avant la correction');
            console.log('   - Les emails avec des problèmes d\'accès COM');
        }
        
        // Test avec un nouvel email fictif pour vérifier le schéma
        console.log('\n🧪 Test d\'insertion d\'un email avec destinataires...');
        
        const testEmailData = {
            outlook_id: `TEST_EMAIL_${Date.now()}`,
            subject: 'Email de test avec destinataires',
            sender_name: 'Test Sender',
            sender_email: 'sender@test.com',
            recipient_email: 'recipient1@test.com; recipient2@test.com',
            received_time: new Date().toISOString(),
            size_kb: 1,
            is_read: false,
            importance: 1,
            folder_name: 'Test Folder',
            category: 'test',
            has_attachment: false
        };
        
        try {
            const result = dbService.saveEmail(testEmailData);
            console.log(`✅ Email de test inséré avec ID: ${result.lastInsertRowid}`);
            
            // Vérifier que l'email a bien été inséré avec les destinataires
            const insertedEmail = dbService.getEmailById(result.lastInsertRowid);
            if (insertedEmail && insertedEmail.recipient_email) {
                console.log(`✅ Destinataires correctement enregistrés: ${insertedEmail.recipient_email}`);
            } else {
                console.log(`❌ Problème: destinataires non enregistrés`);
                console.log('Email inséré:', insertedEmail);
            }
            
            // Nettoyer l'email de test
            dbService.db.prepare('DELETE FROM emails WHERE id = ?').run(result.lastInsertRowid);
            console.log('🧹 Email de test supprimé');
            
        } catch (error) {
            console.error('❌ Erreur lors de l\'insertion de test:', error.message);
        }
        
    } catch (error) {
        console.error('❌ Erreur lors du test:', error);
    }
}

testRecipientEmails();
