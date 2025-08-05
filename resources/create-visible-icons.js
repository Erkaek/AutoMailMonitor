const fs = require('fs');
const path = require('path');

// Créer une icône ICO simple et visible (16x16, 32 bits)
function createSimpleICO() {
    // Données de l'icône 16x16 en format BMP
    // Une enveloppe simple bleue sur fond transparent
    const width = 16;
    const height = 16;
    const bitsPerPixel = 32; // ARGB
    
    // En-tête ICO
    const icoHeader = Buffer.alloc(6);
    icoHeader.writeUInt16LE(0, 0);      // Réservé
    icoHeader.writeUInt16LE(1, 2);      // Type (1 = ICO)
    icoHeader.writeUInt16LE(1, 4);      // Nombre d'images
    
    // Calcul des tailles
    const imageSize = width * height * 4; // 4 bytes par pixel (ARGB)
    const bmpHeaderSize = 40; // Taille de BITMAPINFOHEADER
    const totalImageSize = bmpHeaderSize + imageSize;
    
    // Entrée du répertoire d'icônes
    const dirEntry = Buffer.alloc(16);
    dirEntry.writeUInt8(width, 0);      // Largeur
    dirEntry.writeUInt8(height, 1);     // Hauteur
    dirEntry.writeUInt8(0, 2);          // Couleurs (0 pour >8 bits)
    dirEntry.writeUInt8(0, 3);          // Réservé
    dirEntry.writeUInt16LE(1, 4);       // Plans de couleur
    dirEntry.writeUInt16LE(bitsPerPixel, 6); // Bits par pixel
    dirEntry.writeUInt32LE(totalImageSize, 8); // Taille des données
    dirEntry.writeUInt32LE(22, 12);     // Offset (6 + 16 = 22)
    
    // En-tête BMP (BITMAPINFOHEADER)
    const bmpHeader = Buffer.alloc(40);
    bmpHeader.writeUInt32LE(40, 0);           // Taille de l'en-tête
    bmpHeader.writeInt32LE(width, 4);         // Largeur
    bmpHeader.writeInt32LE(height * 2, 8);    // Hauteur * 2 (pour AND + XOR mask)
    bmpHeader.writeUInt16LE(1, 12);           // Plans
    bmpHeader.writeUInt16LE(bitsPerPixel, 14); // Bits par pixel
    bmpHeader.writeUInt32LE(0, 16);           // Compression (0 = aucune)
    bmpHeader.writeUInt32LE(0, 20);           // Taille image (0 pour non compressé)
    bmpHeader.writeUInt32LE(0, 24);           // X pixels par mètre
    bmpHeader.writeUInt32LE(0, 28);           // Y pixels par mètre
    bmpHeader.writeUInt32LE(0, 32);           // Couleurs utilisées
    bmpHeader.writeUInt32LE(0, 36);           // Couleurs importantes
    
    // Données de pixels - une enveloppe simple
    const pixelData = Buffer.alloc(imageSize);
    
    // Dessiner une enveloppe simple (coordonnées inversées car BMP commence par le bas)
    for (let y = 0; y < height; y++) {
        for (let x = 0; x < width; x++) {
            const index = ((height - 1 - y) * width + x) * 4;
            
            // Dessiner une enveloppe simple
            let alpha = 255;
            let red = 0, green = 0, blue = 0;
            
            // Fond bleu pour l'enveloppe
            if ((y >= 4 && y <= 12) && (x >= 2 && x <= 13)) {
                red = 102;   // #667eea
                green = 126;
                blue = 234;
                alpha = 255;
            }
            // Bord blanc de l'enveloppe
            else if ((y >= 3 && y <= 13) && (x >= 1 && x <= 14) && 
                     (y == 3 || y == 13 || x == 1 || x == 14)) {
                red = 255;
                green = 255;
                blue = 255;
                alpha = 255;
            }
            // Triangle du rabat
            else if (y >= 3 && y <= 8 && x >= 1 && x <= 14) {
                const centerX = 7.5;
                const distFromCenter = Math.abs(x - centerX);
                if (distFromCenter <= (y - 2)) {
                    red = 79;    // Couleur accent #4facfe
                    green = 172;
                    blue = 254;
                    alpha = 255;
                }
            }
            else {
                alpha = 0; // Transparent
            }
            
            // Écrire les données BGRA (ordre inverse)
            pixelData.writeUInt8(blue, index);
            pixelData.writeUInt8(green, index + 1);
            pixelData.writeUInt8(red, index + 2);
            pixelData.writeUInt8(alpha, index + 3);
        }
    }
    
    // Combiner toutes les parties
    return Buffer.concat([icoHeader, dirEntry, bmpHeader, pixelData]);
}

// Créer l'icône
const icoData = createSimpleICO();
const icoPath = path.join(__dirname, 'tray-icon.ico');
fs.writeFileSync(icoPath, icoData);

console.log('Icône tray-icon.ico créée avec une enveloppe bleue visible!');
console.log(`Taille: ${icoData.length} octets`);

// Créer aussi l'icône d'application avec la même méthode mais plus grande
function createAppICO() {
    const width = 32;
    const height = 32;
    const bitsPerPixel = 32;
    
    const icoHeader = Buffer.alloc(6);
    icoHeader.writeUInt16LE(0, 0);
    icoHeader.writeUInt16LE(1, 2);
    icoHeader.writeUInt16LE(1, 4);
    
    const imageSize = width * height * 4;
    const bmpHeaderSize = 40;
    const totalImageSize = bmpHeaderSize + imageSize;
    
    const dirEntry = Buffer.alloc(16);
    dirEntry.writeUInt8(width, 0);
    dirEntry.writeUInt8(height, 1);
    dirEntry.writeUInt8(0, 2);
    dirEntry.writeUInt8(0, 3);
    dirEntry.writeUInt16LE(1, 4);
    dirEntry.writeUInt16LE(bitsPerPixel, 6);
    dirEntry.writeUInt32LE(totalImageSize, 8);
    dirEntry.writeUInt32LE(22, 12);
    
    const bmpHeader = Buffer.alloc(40);
    bmpHeader.writeUInt32LE(40, 0);
    bmpHeader.writeInt32LE(width, 4);
    bmpHeader.writeInt32LE(height * 2, 8);
    bmpHeader.writeUInt16LE(1, 12);
    bmpHeader.writeUInt16LE(bitsPerPixel, 14);
    bmpHeader.writeUInt32LE(0, 16);
    bmpHeader.writeUInt32LE(0, 20);
    bmpHeader.writeUInt32LE(0, 24);
    bmpHeader.writeUInt32LE(0, 28);
    bmpHeader.writeUInt32LE(0, 32);
    bmpHeader.writeUInt32LE(0, 36);
    
    const pixelData = Buffer.alloc(imageSize);
    
    // Dessiner une enveloppe plus détaillée pour 32x32
    for (let y = 0; y < height; y++) {
        for (let x = 0; x < width; x++) {
            const index = ((height - 1 - y) * width + x) * 4;
            
            let alpha = 0;
            let red = 0, green = 0, blue = 0;
            
            // Corps de l'enveloppe
            if ((y >= 8 && y <= 24) && (x >= 4 && x <= 27)) {
                red = 102;
                green = 126;
                blue = 234;
                alpha = 255;
            }
            // Bord blanc
            else if ((y >= 7 && y <= 25) && (x >= 3 && x <= 28) && 
                     (y == 7 || y == 25 || x == 3 || x == 28)) {
                red = 255;
                green = 255;
                blue = 255;
                alpha = 255;
            }
            // Rabat triangulaire
            else if (y >= 7 && y <= 16 && x >= 3 && x <= 28) {
                const centerX = 15.5;
                const distFromCenter = Math.abs(x - centerX);
                if (distFromCenter <= (y - 6) * 1.5) {
                    red = 79;
                    green = 172;
                    blue = 254;
                    alpha = 255;
                }
            }
            // Point de notification
            else if ((x - 22) * (x - 22) + (y - 10) * (y - 10) <= 9) {
                red = 255;
                green = 71;
                blue = 87;
                alpha = 255;
            }
            
            pixelData.writeUInt8(blue, index);
            pixelData.writeUInt8(green, index + 1);
            pixelData.writeUInt8(red, index + 2);
            pixelData.writeUInt8(alpha, index + 3);
        }
    }
    
    return Buffer.concat([icoHeader, dirEntry, bmpHeader, pixelData]);
}

const appIcoData = createAppICO();
const appIcoPath = path.join(__dirname, 'app.ico');
fs.writeFileSync(appIcoPath, appIcoData);

console.log('Icône app.ico créée avec une enveloppe bleue détaillée!');
console.log(`Taille: ${appIcoData.length} octets`);
