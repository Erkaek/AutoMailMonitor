const fs = require('fs');
const path = require('path');

// Créer une icône ICO pour la barre système en utilisant le PNG existant
function createTrayICO() {
    // En-tête ICO (6 octets)
    const header = Buffer.alloc(6);
    header.writeUInt16LE(0, 0);      // Réservé (doit être 0)
    header.writeUInt16LE(1, 2);      // Type (1 pour ICO)
    header.writeUInt16LE(1, 4);      // Nombre d'images

    // Répertoire d'entrée (16 octets par image)
    const dirEntry = Buffer.alloc(16);
    dirEntry.writeUInt8(16, 0);      // Largeur (16 pixels pour tray)
    dirEntry.writeUInt8(16, 1);      // Hauteur (16 pixels pour tray)
    dirEntry.writeUInt8(0, 2);       // Nombre de couleurs (0 pour 24/32 bits)
    dirEntry.writeUInt8(0, 3);       // Réservé
    dirEntry.writeUInt16LE(1, 4);    // Plans de couleur
    dirEntry.writeUInt16LE(32, 6);   // Bits par pixel
    dirEntry.writeUInt32LE(0, 8);    // Taille de l'image (à calculer)
    dirEntry.writeUInt32LE(22, 12);  // Offset vers l'image

    // Image PNG pour la barre système
    const pngPath = path.join(__dirname, 'tray-icon.png');
    const pngData = fs.readFileSync(pngPath);
    
    // Mettre à jour la taille dans le répertoire
    dirEntry.writeUInt32LE(pngData.length, 8);

    // Combiner tous les éléments
    const icoData = Buffer.concat([header, dirEntry, pngData]);
    
    return icoData;
}

// Créer le fichier ICO pour la barre système
const trayIcoData = createTrayICO();
const trayIcoPath = path.join(__dirname, 'tray-icon.ico');
fs.writeFileSync(trayIcoPath, trayIcoData);

console.log('Fichier tray-icon.ico créé avec succès!');
console.log(`Taille: ${trayIcoData.length} octets`);
console.log(`Chemin: ${trayIcoPath}`);
