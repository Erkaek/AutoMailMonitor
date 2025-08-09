const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

console.log('🚀 Construction de l\'exe optimisé...');

try {
  // 1. Nettoyer les anciens builds
  console.log('🧹 Nettoyage des anciens builds...');
  try {
    execSync('npm run clean', { stdio: 'inherit' });
  } catch (error) {
    console.log('ℹ️  Dossier dist déjà propre');
  }

  // 2. S'assurer que les modules natifs sont à jour
  console.log('🔧 Vérification des modules natifs...');
  execSync('npm run rebuild', { stdio: 'inherit' });

  // 3. Builder l'exe avec la configuration optimisée
  console.log('📦 Construction de l\'exe...');
  execSync('npm run build-win', { stdio: 'inherit' });

  // 4. Vérifier que l'exe a été créé
  const distPath = path.join(__dirname, '..', 'dist');
  if (fs.existsSync(distPath)) {
    console.log('✅ Build terminé avec succès !');
    
    // Lister les fichiers créés
    const files = fs.readdirSync(distPath, { withFileTypes: true });
    files.forEach(file => {
      if (file.isFile() && file.name.endsWith('.exe')) {
        const filePath = path.join(distPath, file.name);
        const stats = fs.statSync(filePath);
        const sizeInMB = (stats.size / (1024 * 1024)).toFixed(2);
        console.log(`  🎯 ${file.name} (${sizeInMB} MB)`);
      }
    });
    
    console.log(`\n📂 Fichiers disponibles dans: ${distPath}`);
  } else {
    console.log('❌ Erreur : dossier dist non trouvé');
    process.exit(1);
  }

  console.log('\n🎉 Exe prêt ! Plus de problèmes de modules natifs');

} catch (error) {
  console.error('❌ Erreur lors du build:', error.message);
  console.log('\n💡 Essayez de relancer npm run rebuild puis npm run build-exe');
  process.exit(1);
}
