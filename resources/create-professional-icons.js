const fs = require('fs');
const path = require('path');

// Créer une icône professionnelle avec un design d'enveloppe détaillé
function createProfessionalICO(size = 32) {
    const width = size;
    const height = size;
    
    // En-tête ICO
    const icoHeader = Buffer.alloc(6);
    icoHeader.writeUInt16LE(0, 0);      // Réservé
    icoHeader.writeUInt16LE(1, 2);      // Type ICO
    icoHeader.writeUInt16LE(1, 4);      // Nombre d'images
    
    const imageSize = width * height * 4;
    const bmpHeaderSize = 40;
    const totalImageSize = bmpHeaderSize + imageSize;
    
    // Entrée du répertoire
    const dirEntry = Buffer.alloc(16);
    dirEntry.writeUInt8(width === 256 ? 0 : width, 0);  // 0 pour 256px
    dirEntry.writeUInt8(height === 256 ? 0 : height, 1);
    dirEntry.writeUInt8(0, 2);          // Couleurs
    dirEntry.writeUInt8(0, 3);          // Réservé
    dirEntry.writeUInt16LE(1, 4);       // Plans
    dirEntry.writeUInt16LE(32, 6);      // Bits par pixel
    dirEntry.writeUInt32LE(totalImageSize, 8);
    dirEntry.writeUInt32LE(22, 12);     // Offset
    
    // En-tête BMP
    const bmpHeader = Buffer.alloc(40);
    bmpHeader.writeUInt32LE(40, 0);
    bmpHeader.writeInt32LE(width, 4);
    bmpHeader.writeInt32LE(height * 2, 8);
    bmpHeader.writeUInt16LE(1, 12);
    bmpHeader.writeUInt16LE(32, 14);
    bmpHeader.writeUInt32LE(0, 16);
    bmpHeader.writeUInt32LE(imageSize, 20);
    bmpHeader.writeUInt32LE(0, 24);
    bmpHeader.writeUInt32LE(0, 28);
    bmpHeader.writeUInt32LE(0, 32);
    bmpHeader.writeUInt32LE(0, 36);
    
    // Données de pixels avec design professionnel
    const pixelData = Buffer.alloc(imageSize);
    
    // Couleurs du thème
    const colors = {
        // Enveloppe principale - gradient bleu moderne
        envelope: { r: 99, g: 125, b: 236 },      // #637eec
        envelopeDark: { r: 79, g: 100, b: 200 },  // #4f64c8
        
        // Rabat - bleu plus clair
        flap: { r: 116, g: 140, b: 247 },         // #748cf7
        flapHighlight: { r: 136, g: 160, b: 255 }, // #88a0ff
        
        // Bordures et détails
        border: { r: 255, g: 255, b: 255 },      // Blanc
        shadow: { r: 60, g: 70, b: 120 },        // Ombre
        
        // Point de notification
        notification: { r: 255, g: 87, b: 87 },   // #ff5757
        notificationGlow: { r: 255, g: 120, b: 120 }, // Lueur
        
        // Ligne d'écriture
        lines: { r: 200, g: 210, b: 255 }        // Lignes subtiles
    };
    
    // Facteur d'échelle basé sur la taille
    const scale = size / 32;
    
    for (let y = 0; y < height; y++) {
        for (let x = 0; x < width; x++) {
            const index = ((height - 1 - y) * width + x) * 4;
            
            // Coordonnées normalisées (0-32)
            const nx = x / scale;
            const ny = y / scale;
            
            let color = { r: 0, g: 0, b: 0, a: 0 }; // Transparent par défaut
            
            // Zone principale de l'enveloppe (avec coins arrondis)
            const envelopeLeft = 3;
            const envelopeRight = 29;
            const envelopeTop = 8;
            const envelopeBottom = 24;
            
            // Distance des coins pour l'arrondi
            const cornerRadius = 1.5;
            
            // Fonction pour vérifier si on est dans l'enveloppe arrondie
            function isInEnvelope(x, y) {
                if (x >= envelopeLeft + cornerRadius && x <= envelopeRight - cornerRadius &&
                    y >= envelopeTop && y <= envelopeBottom) return true;
                if (y >= envelopeTop + cornerRadius && y <= envelopeBottom - cornerRadius &&
                    x >= envelopeLeft && x <= envelopeRight) return true;
                
                // Coins arrondis
                const corners = [
                    { cx: envelopeLeft + cornerRadius, cy: envelopeTop + cornerRadius },
                    { cx: envelopeRight - cornerRadius, cy: envelopeTop + cornerRadius },
                    { cx: envelopeLeft + cornerRadius, cy: envelopeBottom - cornerRadius },
                    { cx: envelopeRight - cornerRadius, cy: envelopeBottom - cornerRadius }
                ];
                
                for (let corner of corners) {
                    const dx = x - corner.cx;
                    const dy = y - corner.cy;
                    if (dx * dx + dy * dy <= cornerRadius * cornerRadius) return true;
                }
                return false;
            }
            
            // Corps de l'enveloppe avec gradient
            if (isInEnvelope(nx, ny)) {
                // Gradient vertical
                const gradientFactor = (ny - envelopeTop) / (envelopeBottom - envelopeTop);
                color = {
                    r: Math.round(colors.envelope.r + (colors.envelopeDark.r - colors.envelope.r) * gradientFactor),
                    g: Math.round(colors.envelope.g + (colors.envelopeDark.g - colors.envelope.g) * gradientFactor),
                    b: Math.round(colors.envelope.b + (colors.envelopeDark.b - colors.envelope.b) * gradientFactor),
                    a: 255
                };
            }
            
            // Bordure de l'enveloppe
            const borderThickness = 0.8;
            if (isInEnvelope(nx, ny)) {
                let isBorder = false;
                // Vérifier si on est près du bord
                if (!isInEnvelope(nx - borderThickness, ny) || 
                    !isInEnvelope(nx + borderThickness, ny) ||
                    !isInEnvelope(nx, ny - borderThickness) || 
                    !isInEnvelope(nx, ny + borderThickness)) {
                    isBorder = true;
                }
                
                if (isBorder) {
                    color = { ...colors.border, a: 255 };
                }
            }
            
            // Rabat triangulaire (partie supérieure)
            const flapTop = 6;
            const flapBottom = 16;
            const centerX = 16;
            
            if (ny >= flapTop && ny <= flapBottom && nx >= envelopeLeft && nx <= envelopeRight) {
                const flapProgress = (ny - flapTop) / (flapBottom - flapTop);
                const flapWidth = (envelopeRight - envelopeLeft) * (1 - flapProgress * 0.7);
                const flapLeft = centerX - flapWidth / 2;
                const flapRight = centerX + flapWidth / 2;
                
                if (nx >= flapLeft && nx <= flapRight) {
                    // Gradient pour le rabat
                    const flapGradient = 1 - flapProgress;
                    color = {
                        r: Math.round(colors.flap.r + (colors.flapHighlight.r - colors.flap.r) * flapGradient),
                        g: Math.round(colors.flap.g + (colors.flapHighlight.g - colors.flap.g) * flapGradient),
                        b: Math.round(colors.flap.b + (colors.flapHighlight.b - colors.flap.b) * flapGradient),
                        a: 255
                    };
                }
            }
            
            // Lignes de texte à l'intérieur
            if (size >= 32) {
                for (let line = 0; line < 3; line++) {
                    const lineY = 12 + line * 2.5;
                    const lineLeft = 6;
                    const lineRight = 26 - line * 1.5; // Lignes de longueur décroissante
                    
                    if (Math.abs(ny - lineY) <= 0.3 && nx >= lineLeft && nx <= lineRight) {
                        color = { ...colors.lines, a: 180 };
                    }
                }
            }
            
            // Point de notification (coin supérieur droit)
            const notifX = 24;
            const notifY = 9;
            const notifRadius = size >= 32 ? 2.5 : 1.5;
            const distToNotif = Math.sqrt((nx - notifX) * (nx - notifX) + (ny - notifY) * (ny - notifY));
            
            if (distToNotif <= notifRadius) {
                if (distToNotif <= notifRadius * 0.7) {
                    color = { ...colors.notification, a: 255 };
                } else {
                    // Lueur autour de la notification
                    const glowFactor = 1 - (distToNotif - notifRadius * 0.7) / (notifRadius * 0.3);
                    color = {
                        r: Math.round(colors.notificationGlow.r),
                        g: Math.round(colors.notificationGlow.g),
                        b: Math.round(colors.notificationGlow.b),
                        a: Math.round(255 * glowFactor)
                    };
                }
            }
            
            // Ombre portée (effet de profondeur)
            if (size >= 32) {
                const shadowOffsetX = 1;
                const shadowOffsetY = 1;
                if (isInEnvelope(nx - shadowOffsetX, ny - shadowOffsetY) && 
                    color.a === 0) {
                    color = { ...colors.shadow, a: 60 };
                }
            }
            
            // Écrire les données BGRA
            pixelData.writeUInt8(color.b, index);
            pixelData.writeUInt8(color.g, index + 1);
            pixelData.writeUInt8(color.r, index + 2);
            pixelData.writeUInt8(color.a, index + 3);
        }
    }
    
    return Buffer.concat([icoHeader, dirEntry, bmpHeader, pixelData]);
}

// Créer les différentes tailles d'icônes
const sizes = [16, 32, 48];

for (let size of sizes) {
    const icoData = createProfessionalICO(size);
    
    if (size === 16) {
        fs.writeFileSync(path.join(__dirname, 'tray-icon.ico'), icoData);
        console.log(`✨ tray-icon.ico (${size}x${size}) - Icône système professionnelle créée`);
    } else if (size === 32) {
        fs.writeFileSync(path.join(__dirname, 'app.ico'), icoData);
        console.log(`✨ app.ico (${size}x${size}) - Icône application professionnelle créée`);
    } else {
        fs.writeFileSync(path.join(__dirname, `app-${size}.ico`), icoData);
        console.log(`✨ app-${size}.ico (${size}x${size}) - Icône haute résolution créée`);
    }
}

console.log('\n🎨 Icônes professionnelles créées avec :');
console.log('• Design d\'enveloppe moderne avec coins arrondis');
console.log('• Gradient bleu élégant');
console.log('• Rabat triangulaire avec éclairage');
console.log('• Lignes de texte subtiles');
console.log('• Point de notification rouge avec lueur');
console.log('• Ombre portée pour la profondeur');
console.log('• Bordures blanches nettes');
console.log('\n🚀 Ces icônes sont beaucoup plus professionnelles !');
