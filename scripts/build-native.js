const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

console.log('🔧 Optimisation des modules natifs pour Electron...');

try {
  // Nettoyer le cache d'Electron
  console.log('📝 Nettoyage du cache Electron...');
  try {
    execSync('npx electron-builder install-app-deps', { stdio: 'inherit' });
  } catch (error) {
    console.log('ℹ️  Cache déjà propre ou commande non disponible');
  }

  // Recompiler better-sqlite3 pour Electron
  console.log('🔨 Recompilation de better-sqlite3 pour Electron...');
  execSync('npx electron-rebuild -f -w better-sqlite3', { stdio: 'inherit' });

  // Vérifier que les fichiers natifs sont présents
  const nativeModulePath = path.join(__dirname, '..', 'node_modules', 'better-sqlite3', 'build');
  if (fs.existsSync(nativeModulePath)) {
    console.log('✅ Modules natifs compilés avec succès');
    
    // Lister les fichiers compilés
    const files = fs.readdirSync(path.join(nativeModulePath, 'Release'), { withFileTypes: true });
    files.forEach(file => {
      if (file.isFile()) {
        console.log(`  📄 ${file.name}`);
      }
    });
  } else {
    console.log('❌ Erreur : modules natifs non trouvés');
    process.exit(1);
  }

  console.log('🎉 Compilation terminée ! Modules prêts pour l\'exe');

} catch (error) {
  console.error('❌ Erreur lors de la compilation:', error.message);
  console.log('\n💡 Solution alternative : installez Windows Build Tools si pas déjà fait');
  console.log('   npm install --global windows-build-tools');
  process.exit(1);
}
