// Création d'une icône simple pour la barre système
const fs = require('fs');
const path = require('path');

// Image PNG simple 32x32 avec un fond bleu et un enveloppe blanche
const base64Icon = `iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAMiSURBVFhH7ZdLaxtHFIafM5JsyZJlSZZkOZIdO3Ecx3mR4jhxHCdOnMRJnCRx4iROnMRJnDiJkyROnMRJ3NiJncRJnMRJnMSNmzhx4iZu3MRNnLiJE7epnbi+9W95/+eduXPmzOys5DhO8H/jfwCUlJSUlJSUlJSUlJSUlJSUlJSUlJSUlJSUlJSUlJSUlJSUlJSUlJSUlJSUlJSUlJSUlJSUlJSUlJSUlJSUlJSUlJSUlJT+D/gfBEpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSv8P/A8C/wGhQziH5w6Z6wAAAABJRU5ErkJggg==`;

// Convertir base64 en buffer et sauvegarder
const iconBuffer = Buffer.from(base64Icon, 'base64');
const iconPath = path.join(__dirname, 'tray-icon.png');

fs.writeFileSync(iconPath, iconBuffer);
console.log('Icône PNG créée:', iconPath);
