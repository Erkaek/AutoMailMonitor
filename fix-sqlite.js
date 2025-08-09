#!/usr/bin/env node

/**
 * Script de diagnostic et réparation automatique pour better-sqlite3
 * Résout automatiquement les problèmes de version NODE_MODULE_VERSION
 */

const { execSync } = require('child_process');
const path = require('path');
const fs = require('fs');

console.log('🔧 Diagnostic et réparation automatique better-sqlite3...\n');

// 1. Vérifier la version de Node.js
try {
  const nodeVersion = process.version;
  console.log(`✅ Version Node.js: ${nodeVersion}`);
} catch (error) {
  console.error('❌ Impossible de détecter la version Node.js');
  process.exit(1);
}

// 2. Vérifier si better-sqlite3 est installé
const sqlitePath = path.join(__dirname, 'node_modules', 'better-sqlite3');
if (!fs.existsSync(sqlitePath)) {
  console.log('⚠️  better-sqlite3 n\'est pas installé');
  process.exit(1);
}

// 3. Tenter de charger better-sqlite3
let needsRebuild = false;
try {
  require('better-sqlite3');
  console.log('✅ better-sqlite3 fonctionne correctement');
} catch (error) {
  if (error.message.includes('NODE_MODULE_VERSION')) {
    console.log('❌ Problème de version détecté:', error.message.split('\n')[0]);
    needsRebuild = true;
  } else {
    console.error('❌ Erreur inattendue:', error.message);
    process.exit(1);
  }
}

// 4. Rebuild automatique si nécessaire
if (needsRebuild) {
  console.log('\n🔨 Rebuild automatique en cours...');
  try {
    execSync('npm rebuild better-sqlite3', { 
      stdio: 'inherit',
      cwd: __dirname 
    });
    console.log('✅ Rebuild terminé avec succès');
    
    // Test final
    require('better-sqlite3');
    console.log('✅ better-sqlite3 fonctionne maintenant correctement');
  } catch (rebuildError) {
    console.error('❌ Échec du rebuild:', rebuildError.message);
    console.log('\n💡 Solutions alternatives:');
    console.log('   1. npm install --force');
    console.log('   2. Supprimer node_modules et reinstaller');
    console.log('   3. Utiliser une version différente de Node.js');
    process.exit(1);
  }
} else {
  console.log('✅ Aucune action nécessaire');
}

console.log('\n🎉 Diagnostic terminé - better-sqlite3 est prêt !');
