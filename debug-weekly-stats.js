/**
 * Script de debug pour vérifier l'état des statistiques hebdomadaires
 */

const Database = require('better-sqlite3');
const path = require('path');

function debugWeeklyStats() {
    const dbPath = path.join(__dirname, 'emails.db');
    const db = new Database(dbPath);
    
    try {
        console.log('🔍 DIAGNOSTIC: Statistiques hebdomadaires\n');
        
        // 1. Vérifier les tables existantes
        console.log('📋 Tables existantes:');
        const tables = db.prepare("SELECT name FROM sqlite_master WHERE type='table'").all();
        tables.forEach(table => {
            console.log(`  - ${table.name}`);
        });
        
        // 2. Vérifier si weekly_stats existe
        const weeklyStatsExists = tables.find(t => t.name === 'weekly_stats');
        console.log(`\n📊 Table weekly_stats: ${weeklyStatsExists ? '✅ Existe' : '❌ N\'existe pas'}`);
        
        if (weeklyStatsExists) {
            // 3. Compter les lignes
            const count = db.prepare('SELECT COUNT(*) as count FROM weekly_stats').get();
            console.log(`📈 Nombre de lignes: ${count.count}`);
            
            // 4. Afficher la structure
            console.log('\n🏗️ Structure de la table:');
            const pragma = db.prepare('PRAGMA table_info(weekly_stats)').all();
            pragma.forEach(col => {
                console.log(`  - ${col.name}: ${col.type} ${col.notnull ? 'NOT NULL' : ''} ${col.pk ? 'PRIMARY KEY' : ''}`);
            });
            
            // 5. Échantillon de données
            if (count.count > 0) {
                console.log('\n📄 Échantillon de données:');
                const sample = db.prepare('SELECT * FROM weekly_stats LIMIT 3').all();
                sample.forEach((row, i) => {
                    console.log(`\n  Ligne ${i + 1}:`);
                    Object.keys(row).forEach(key => {
                        console.log(`    ${key}: ${row[key]}`);
                    });
                });
            } else {
                console.log('\n📄 Aucune donnée dans weekly_stats');
            }
        }
        
        // 6. Vérifier les emails pour générer des stats
        console.log('\n📧 Statistiques emails par folder_name:');
        const emailStats = db.prepare(`
            SELECT 
                folder_name,
                COUNT(*) as total,
                SUM(CASE WHEN is_read = 0 THEN 1 ELSE 0 END) as unread,
                DATE(received_time) as date
            FROM emails 
            GROUP BY folder_name, DATE(received_time)
            ORDER BY received_time DESC
            LIMIT 10
        `).all();
        
        if (emailStats.length > 0) {
            emailStats.forEach(stat => {
                console.log(`  📁 ${stat.folder_name}: ${stat.total} emails (${stat.unread} non lus) - ${stat.date}`);
            });
        } else {
            console.log('  ❌ Aucun email trouvé');
        }
        
        // 7. Test de l'identifiant de semaine actuel
        console.log('\n📅 Test identifiant semaine ISO:');
        const now = new Date();
        const tempDate = new Date(now);
        tempDate.setHours(0, 0, 0, 0);
        tempDate.setDate(tempDate.getDate() + 3 - (tempDate.getDay() + 6) % 7);
        const week1 = new Date(tempDate.getFullYear(), 0, 4);
        const weekNumber = 1 + Math.round(((tempDate.getTime() - week1.getTime()) / 86400000 - 3 + (week1.getDay() + 6) % 7) / 7);
        const year = tempDate.getFullYear();
        const identifier = `S${weekNumber}-${year}`;
        
        console.log(`  Semaine actuelle: ${identifier}`);
        console.log(`  Numéro: ${weekNumber}, Année: ${year}`);
        
        // 8. Test de création manuelle de weekly_stats si elle n'existe pas
        if (!weeklyStatsExists) {
            console.log('\n🔧 Création de la table weekly_stats...');
            db.exec(`
                CREATE TABLE IF NOT EXISTS weekly_stats (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    week_identifier TEXT NOT NULL,
                    week_number INTEGER NOT NULL,
                    week_year INTEGER NOT NULL,
                    week_start_date TEXT NOT NULL,
                    week_end_date TEXT NOT NULL,
                    folder_type TEXT NOT NULL DEFAULT 'Mails simples',
                    emails_received INTEGER DEFAULT 0,
                    emails_treated INTEGER DEFAULT 0,
                    manual_adjustments INTEGER DEFAULT 0,
                    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(week_identifier, folder_type)
                )
            `);
            console.log('✅ Table weekly_stats créée');
            
            // Insérer une ligne test pour la semaine actuelle
            const monday = new Date(tempDate);
            monday.setDate(monday.getDate() - (monday.getDay() + 6) % 7);
            const sunday = new Date(monday);
            sunday.setDate(sunday.getDate() + 6);
            
            db.prepare(`
                INSERT OR REPLACE INTO weekly_stats 
                (week_identifier, week_number, week_year, week_start_date, week_end_date, folder_type, emails_received)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            `).run(identifier, weekNumber, year, monday.toISOString().split('T')[0], sunday.toISOString().split('T')[0], 'Mails simples', 14);
            
            console.log('✅ Ligne test ajoutée');
        }
        
    } catch (error) {
        console.error('❌ Erreur:', error);
    } finally {
        db.close();
    }
}

// Exécuter le diagnostic
debugWeeklyStats();
