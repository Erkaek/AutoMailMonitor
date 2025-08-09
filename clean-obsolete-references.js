const fs = require('fs');
const path = require('path');

console.log('🧹 NETTOYAGE DES RÉFÉRENCES OBSOLÈTES');
console.log('====================================');

// Fichiers à analyser et corriger
const filesToCheck = [
    'src/services/optimizedDatabaseService.js',
    'src/services/unifiedMonitoringService.js',
    'src/main/index.js'
];

// Patterns obsolètes à corriger
const obsoletePatterns = [
    {
        pattern: /sender_name/g,
        replacement: 'sender_email',
        description: 'sender_name → sender_email'
    },
    {
        pattern: /entry_id/g,
        replacement: 'outlook_id',
        description: 'entry_id → outlook_id'
    },
    {
        pattern: /sent_time/g,
        replacement: 'received_time',
        description: 'sent_time → received_time'
    },
    {
        pattern: /treated_time/g,
        replacement: 'deleted_at',
        description: 'treated_time → deleted_at'
    },
    {
        pattern: /folder_path/g,
        replacement: 'folder_name',
        description: 'folder_path → folder_name'
    },
    {
        pattern: /recipient_email/g,
        replacement: '/* recipient_email REMOVED */',
        description: 'recipient_email → REMOVED'
    },
    {
        pattern: /is_replied/g,
        replacement: 'is_treated',
        description: 'is_replied → is_treated'
    },
    {
        pattern: /has_attachment/g,
        replacement: '/* has_attachment REMOVED */',
        description: 'has_attachment → REMOVED'
    },
    {
        pattern: /body_preview/g,
        replacement: '/* body_preview REMOVED */',
        description: 'body_preview → REMOVED'
    },
    {
        pattern: /importance/g,
        replacement: '/* importance REMOVED */',
        description: 'importance → REMOVED'
    },
    {
        pattern: /size_bytes/g,
        replacement: '/* size_bytes REMOVED */',
        description: 'size_bytes → REMOVED'
    }
];

let totalReplacements = 0;

for (const filePath of filesToCheck) {
    const fullPath = path.resolve(filePath);
    
    if (!fs.existsSync(fullPath)) {
        console.log(`⚠️  Fichier non trouvé: ${filePath}`);
        continue;
    }
    
    console.log(`\n📁 Traitement: ${filePath}`);
    
    let content = fs.readFileSync(fullPath, 'utf8');
    let fileReplacements = 0;
    
    for (const {pattern, replacement, description} of obsoletePatterns) {
        const matches = content.match(pattern);
        if (matches) {
            console.log(`   🔧 ${description}: ${matches.length} occurrences`);
            content = content.replace(pattern, replacement);
            fileReplacements += matches.length;
        }
    }
    
    if (fileReplacements > 0) {
        // Sauvegarder le fichier modifié
        fs.writeFileSync(fullPath, content);
        console.log(`   ✅ ${fileReplacements} corrections appliquées`);
        totalReplacements += fileReplacements;
    } else {
        console.log(`   ✅ Aucune correction nécessaire`);
    }
}

console.log(`\n🎯 RÉSUMÉ: ${totalReplacements} corrections totales appliquées`);

if (totalReplacements > 0) {
    console.log('\n⚠️  ATTENTION: Vérifiez manuellement que les remplacements sont corrects !');
    console.log('   Certains remplacements peuvent nécessiter des ajustements contextuels.');
}

console.log('\n✅ NETTOYAGE TERMINÉ');
