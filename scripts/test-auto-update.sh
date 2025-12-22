#!/bin/bash
# Test rapide du système de mise à jour automatique
# Ce script configure un environnement de test local

set -e

echo "🧪 Configuration du test de mise à jour automatique"
echo ""

# 1. Créer le dossier de test
echo "📁 Création du dossier de test..."
TEST_DIR="/tmp/autoupdater-test"
rm -rf "$TEST_DIR"
mkdir -p "$TEST_DIR"
cd "$TEST_DIR"

# 2. Créer le fichier latest.yml
echo "📝 Création de latest.yml..."
cat > latest.yml << 'EOF'
version: 1.0.1
releaseDate: 2024-01-15T10:00:00.000Z
files:
  - url: AutoMailMonitor-Setup-1.0.1.exe
    sha512: YmFzZTY0IGVuY29kZWQgc2hhNTEyIGhhc2g=
    size: 50000000
path: AutoMailMonitor-Setup-1.0.1.exe
sha512: YmFzZTY0IGVuY29kZWQgc2hhNTEyIGhhc2g=
releaseNotes: |
  ## Nouveautés v1.0.1
  
  ### Ajouté
  - Système de logs amélioré avec filtres
  - Gestionnaire de mises à jour robuste
  - Notifications UI modernes
  
  ### Corrigé
  - Historique des emails manquant
  - Erreurs JavaScript
  - Erreurs PowerShell
  
  ### Performance
  - Retry automatique avec backoff
  - Timeout protection (30s)
  - Logging détaillé
EOF

# 3. Créer un fichier exe factice
echo "📦 Création du fichier exe factice..."
dd if=/dev/zero of=AutoMailMonitor-Setup-1.0.1.exe bs=1M count=10 2>/dev/null

# 4. Vérifier que http-server est installé
if ! command -v http-server &> /dev/null; then
    echo "⚠️  http-server n'est pas installé."
    echo "📥 Installation via npm..."
    npm install -g http-server
fi

# 5. Démarrer le serveur
echo ""
echo "✅ Configuration terminée !"
echo ""
echo "📂 Dossier de test : $TEST_DIR"
echo "📄 Fichiers créés :"
ls -lh "$TEST_DIR"
echo ""
echo "🚀 Pour démarrer le serveur de test :"
echo "   cd $TEST_DIR"
echo "   http-server -p 8080 --cors"
echo ""
echo "🔧 Pour tester l'application :"
echo "   1. Modifier package.json : \"version\": \"1.0.0\""
echo "   2. Lancer l'app : npm start"
echo "   3. Observer les logs dans l'onglet Logs"
echo ""
echo "🌐 URL du serveur : http://localhost:8080"
echo "📋 latest.yml : http://localhost:8080/latest.yml"
echo ""
echo "💡 Conseil : Ouvrir 2 terminaux"
echo "   Terminal 1 : http-server (dans $TEST_DIR)"
echo "   Terminal 2 : npm start (dans le projet)"
echo ""

# Option pour démarrer directement le serveur
read -p "Démarrer le serveur maintenant ? (y/N) " -n 1 -r
echo
if [[ $REPLY =~ ^[Yy]$ ]]; then
    echo ""
    echo "🌐 Démarrage du serveur sur http://localhost:8080"
    echo "   Appuyez sur Ctrl+C pour arrêter"
    echo ""
    http-server -p 8080 --cors
else
    echo ""
    echo "✋ Serveur non démarré. Lancez-le manuellement avec :"
    echo "   cd $TEST_DIR && http-server -p 8080 --cors"
    echo ""
fi
