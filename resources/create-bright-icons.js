const fs = require('fs');
const path = require('path');

// Créer un ICO avec des couleurs vives et un design simple mais visible
function createBrightICO() {
    // Créer une icône 16x16 avec une enveloppe orange/rouge vive
    const width = 16;
    const height = 16;
    
    // En-tête ICO
    const icoHeader = Buffer.alloc(6);
    icoHeader.writeUInt16LE(0, 0);      // Réservé
    icoHeader.writeUInt16LE(1, 2);      // Type ICO
    icoHeader.writeUInt16LE(1, 4);      // Nombre d'images
    
    // Entrée du répertoire
    const dirEntry = Buffer.alloc(16);
    dirEntry.writeUInt8(width, 0);      // Largeur
    dirEntry.writeUInt8(height, 1);     // Hauteur
    dirEntry.writeUInt8(0, 2);          // Couleurs
    dirEntry.writeUInt8(0, 3);          // Réservé
    dirEntry.writeUInt16LE(1, 4);       // Plans
    dirEntry.writeUInt16LE(32, 6);      // Bits par pixel
    dirEntry.writeUInt32LE(1064, 8);    // Taille des données
    dirEntry.writeUInt32LE(22, 12);     // Offset
    
    // En-tête BMP
    const bmpHeader = Buffer.alloc(40);
    bmpHeader.writeUInt32LE(40, 0);           // Taille en-tête
    bmpHeader.writeInt32LE(width, 4);         // Largeur
    bmpHeader.writeInt32LE(height * 2, 8);    // Hauteur (avec masque)
    bmpHeader.writeUInt16LE(1, 12);           // Plans
    bmpHeader.writeUInt16LE(32, 14);          // Bits par pixel
    bmpHeader.writeUInt32LE(0, 16);           // Compression
    bmpHeader.writeUInt32LE(1024, 20);        // Taille image
    bmpHeader.writeUInt32LE(0, 24);           // X pixels/m
    bmpHeader.writeUInt32LE(0, 28);           // Y pixels/m
    bmpHeader.writeUInt32LE(0, 32);           // Couleurs utilisées
    bmpHeader.writeUInt32LE(0, 36);           // Couleurs importantes
    
    // Données de pixels - enveloppe orange vive
    const pixelData = Buffer.alloc(1024); // 16x16x4 = 1024 bytes
    
    for (let y = 0; y < height; y++) {
        for (let x = 0; x < width; x++) {
            const index = ((height - 1 - y) * width + x) * 4;
            
            let blue = 0, green = 0, red = 0, alpha = 0;
            
            // Enveloppe orange/rouge vive
            if (y >= 4 && y <= 11 && x >= 2 && x <= 13) {
                // Corps orange vif
                blue = 0;
                green = 165;
                red = 255;
                alpha = 255;
            }
            // Bordure blanche
            else if (y >= 3 && y <= 12 && x >= 1 && x <= 14 && 
                     (y == 3 || y == 12 || x == 1 || x == 14)) {
                blue = 255;
                green = 255;
                red = 255;
                alpha = 255;
            }
            // Rabat rouge vif
            else if (y >= 3 && y <= 7 && x >= 1 && x <= 14) {
                const centerX = 7.5;
                const distFromCenter = Math.abs(x - centerX);
                if (distFromCenter <= (y - 2)) {
                    blue = 0;
                    green = 0;
                    red = 255;
                    alpha = 255;
                }
            }
            
            // Écrire BGRA
            pixelData.writeUInt8(blue, index);
            pixelData.writeUInt8(green, index + 1);
            pixelData.writeUInt8(red, index + 2);
            pixelData.writeUInt8(alpha, index + 3);
        }
    }
    
    return Buffer.concat([icoHeader, dirEntry, bmpHeader, pixelData]);
}

// Créer l'icône pour la barre système
const brightIco = createBrightICO();
fs.writeFileSync(path.join(__dirname, 'tray-icon.ico'), brightIco);

// Créer aussi l'icône d'application (32x32) avec les mêmes couleurs vives
function createBrightAppICO() {
    const width = 32;
    const height = 32;
    
    const icoHeader = Buffer.alloc(6);
    icoHeader.writeUInt16LE(0, 0);
    icoHeader.writeUInt16LE(1, 2);
    icoHeader.writeUInt16LE(1, 4);
    
    const dirEntry = Buffer.alloc(16);
    dirEntry.writeUInt8(width, 0);
    dirEntry.writeUInt8(height, 1);
    dirEntry.writeUInt8(0, 2);
    dirEntry.writeUInt8(0, 3);
    dirEntry.writeUInt16LE(1, 4);
    dirEntry.writeUInt16LE(32, 6);
    dirEntry.writeUInt32LE(4136, 8);    // 40 + 4096
    dirEntry.writeUInt32LE(22, 12);
    
    const bmpHeader = Buffer.alloc(40);
    bmpHeader.writeUInt32LE(40, 0);
    bmpHeader.writeInt32LE(width, 4);
    bmpHeader.writeInt32LE(height * 2, 8);
    bmpHeader.writeUInt16LE(1, 12);
    bmpHeader.writeUInt16LE(32, 14);
    bmpHeader.writeUInt32LE(0, 16);
    bmpHeader.writeUInt32LE(4096, 20);
    bmpHeader.writeUInt32LE(0, 24);
    bmpHeader.writeUInt32LE(0, 28);
    bmpHeader.writeUInt32LE(0, 32);
    bmpHeader.writeUInt32LE(0, 36);
    
    const pixelData = Buffer.alloc(4096); // 32x32x4
    
    for (let y = 0; y < height; y++) {
        for (let x = 0; x < width; x++) {
            const index = ((height - 1 - y) * width + x) * 4;
            
            let blue = 0, green = 0, red = 0, alpha = 0;
            
            // Corps orange vif plus grand
            if (y >= 8 && y <= 23 && x >= 4 && x <= 27) {
                blue = 0;
                green = 165;
                red = 255;
                alpha = 255;
            }
            // Bordure blanche
            else if (y >= 7 && y <= 24 && x >= 3 && x <= 28 && 
                     (y == 7 || y == 24 || x == 3 || x == 28)) {
                blue = 255;
                green = 255;
                red = 255;
                alpha = 255;
            }
            // Rabat rouge vif
            else if (y >= 7 && y <= 15 && x >= 3 && x <= 28) {
                const centerX = 15.5;
                const distFromCenter = Math.abs(x - centerX);
                if (distFromCenter <= (y - 6) * 1.2) {
                    blue = 0;
                    green = 0;
                    red = 255;
                    alpha = 255;
                }
            }
            // Point de notification jaune vif
            else if ((x - 23) * (x - 23) + (y - 9) * (y - 9) <= 12) {
                blue = 0;
                green = 255;
                red = 255;
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

const brightAppIco = createBrightAppICO();
fs.writeFileSync(path.join(__dirname, 'app.ico'), brightAppIco);

console.log('🎨 Icônes orange/rouge vives créées !');
console.log('- tray-icon.ico: Enveloppe orange 16x16');
console.log('- app.ico: Enveloppe orange 32x32 avec notification');
console.log('Ces icônes devraient être beaucoup plus visibles !');
