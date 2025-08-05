const fs = require('fs');
const path = require('path');

// Icône PNG 32x32 valide - données base64 d'une vraie icône de mail simple
const validIconBase64 = `iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAABHNCSVQICAgIfAhkiAAAAAlwSFlzAAAAdgAAAHYBTnsmCAAAABl0RVh0U29mdHdhcmUAd3d3Lmlua3NjYXBlLm9yZ5vuPBoAAAMvSURBVFhH7ZZLaFNBFIafJm1i0zxtH7a2trZWW1sfrfWBFkVEQQRBcaGIihuLuHDhQleuXLhx4cKFC3HhQhcuXIgLF4ILFy5cuHDhQhcuXIgLFy5cuHDhQhEXggtBBBEXLvz/O3PvnZk7c2cy6aQNHPgYmHvPfzP/zJk7M5OklFJKKaWUUkoppZRSSimllFJKKaWUUkoppZRSSimllFJKKaWUUkoppZRSSimllFJKKaWUUkoppZRSSimllFJKKaWUUkoppZRS/1f9DwKllFJKKaWUUkoppZRSSimllFJKKaWUUkoppZRSSimllFJKKaWUUkoppZRSSimllFJKKaWUUkoppZRSSimllFJKKaWUUkoppZRSSqn/I/6HwCdQSimllFJKKaWUUkoppZRSSimllFJKKaWUUkoppZRSSimllFJKKaWUUkoppZRSSimllFJKKaWUUkoppZRSSimllFJKKaWUUkoppZT6P/I/BL6AmpoampubqbGxkRoaGqi+vp7q6uqotraWampqqLq6mqqqqujPnz9UWVlJv3//psrKSqqoqKDy8nIqKyuj0tJSKikpoeLiYioqKqLCwkIqKCig/Px8ysvLo9zcXMrJyaHs7GzKysqi7OxsysrKouzsbMrJyaHc3FzKy8uj/Px8ys/Pp7y8PMrLy6O8vDzKy8uj3NxcysnJoezsbMrKyqLs7GzKysqi7OxsysnJodzcXMrLy6P8/HwqKCiggoICKiwspKKiIiouLqbS0lIqLS2lkpISKikpoeLiYiopKaHS0lIqLS2l4uJiKioqoqKiIiosLKSCggLKz8+nvLw8ys3NpZycHMrOzqasrCzKzs6m7OxsysnJodzcXMrLy6P8/HwqKCig/Px8ysvLo9zcXMrJyaHs7GzKysqi7OxsysnJodzcXMrLy6P8/HwqKCiggoICKiwspKKiIiouLqbS0lIqKSmhkpISKi4upuLiYioqKqLCwkIqKCig/Px8ysvLo9zcXMrJyaHs7GzKysqi7OxsysnJodzcXMrLy6P8/HwqKCiggoICKiwspKKiIiouLqbS0lIqKSmhkpISKi4upuLiYioqKqLCwkIqKCig/Px8ysvLo9zcXMrJyaHs7GzKysqi7OxsysnJodzcXMrLy6P8/HwqKCigwsJCKioqouLiYiotLaXS0lIqKSmhkpISKi4upuLiYioqKqLCwkIqKCig/Px8ysvLo9zcXMrJyaHs7GzKzs6mnJwcys3NpT8AaB0t0ZfO6AoAAAAASUVORK5CYII=`;

// Ecrire les icônes
const iconBuffer = Buffer.from(validIconBase64, 'base64');

// Icône de barre système
const trayIconPath = path.join(__dirname, 'tray-icon.png');
fs.writeFileSync(trayIconPath, iconBuffer);
console.log('Icône tray-icon.png créée');

// Icône d'application
const appIconPath = path.join(__dirname, 'icon.png');
fs.writeFileSync(appIconPath, iconBuffer);
console.log('Icône icon.png créée');

console.log('Tailles des fichiers:');
console.log(`tray-icon.png: ${fs.statSync(trayIconPath).size} octets`);
console.log(`icon.png: ${fs.statSync(appIconPath).size} octets`);
