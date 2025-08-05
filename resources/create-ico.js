const fs = require('fs');
const path = require('path');

// Créer une icône ICO simple en utilisant la structure binaire ICO
function createICOFile() {
    // En-tête ICO (6 octets)
    const header = Buffer.alloc(6);
    header.writeUInt16LE(0, 0);      // Réservé (doit être 0)
    header.writeUInt16LE(1, 2);      // Type (1 pour ICO)
    header.writeUInt16LE(1, 4);      // Nombre d'images

    // Répertoire d'entrée (16 octets par image)
    const dirEntry = Buffer.alloc(16);
    dirEntry.writeUInt8(32, 0);      // Largeur (32 pixels)
    dirEntry.writeUInt8(32, 1);      // Hauteur (32 pixels)
    dirEntry.writeUInt8(0, 2);       // Nombre de couleurs (0 pour 24/32 bits)
    dirEntry.writeUInt8(0, 3);       // Réservé
    dirEntry.writeUInt16LE(1, 4);    // Plans de couleur
    dirEntry.writeUInt16LE(32, 6);   // Bits par pixel
    dirEntry.writeUInt32LE(0, 8);    // Taille de l'image (à calculer)
    dirEntry.writeUInt32LE(22, 12);  // Offset vers l'image

    // Image PNG intégrée (utilisons notre PNG existant)
    const pngPath = path.join(__dirname, 'icon.png');
    const pngData = fs.readFileSync(pngPath);
    
    // Mettre à jour la taille dans le répertoire
    dirEntry.writeUInt32LE(pngData.length, 8);

    // Combiner tous les éléments
    const icoData = Buffer.concat([header, dirEntry, pngData]);
    
    return icoData;
}

// Créer le fichier ICO
const icoData = createICOFile();
const icoPath = path.join(__dirname, 'app.ico');
fs.writeFileSync(icoPath, icoData);

console.log('Fichier app.ico créé avec succès!');
console.log(`Taille: ${icoData.length} octets`);
console.log(`Chemin: ${icoPath}`);
